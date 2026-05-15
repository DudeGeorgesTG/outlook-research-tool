# Outlook Account Research Tool

Python-based Microsoft account research utility for educational and authorized security testing purposes.

## Features

* Microsoft account OAuth authentication
* Profile information retrieval
* Mailbox metadata extraction
* Inbox statistics collection
* Linked service detection
* JSON result exporting
* UTF-8 output support
* Config-driven service matching

---

## Preview

```json
{
  "success": true,
  "email": "example@hotmail.com",
  "profile": {
    "name": "John Doe",
    "country": "US",
    "flag": "🇺🇸"
  },
  "mailbox": {
    "timezone": "UTC",
    "total_emails": 1200,
    "unread_emails": 32
  }
}
```

---

## Requirements

* Python 3.9+
* requests

Install dependencies:

```bash
pip install requests
```

---

## Project Structure

```bash
.
├── main.py
├── config.json
├── valid_accounts.txt
├── account_details.json
└── README.md
```

---

## Configuration

Example `config.json`:

```json
{
  "services": {
    "Netflix": {
      "sender": "netflix",
      "category": "streaming"
    },
    "PayPal": {
      "sender": "paypal",
      "category": "finance"
    },
    "Amazon": {
      "sender": "amazon",
      "category": "shopping"
    }
  }
}
```

---

## Usage

Run the script:

```bash
python main.py
```

Input format:

```bash
email@example.com:password
```

---

## Output Files

### `valid_accounts.txt`

Stores successfully authenticated accounts.

Example:

```txt
email@example.com:password
```

### `account_details.json`

Stores structured account data and mailbox information.

---

## Extracted Information

### Profile Information

* Full name
* First/last name
* Country
* Country flag
* Birthdate
* Phone number
* Email verification status

### Mailbox Information

* Timezone
* Time format
* Date format
* Dark mode status
* Focused inbox status
* Total emails
* Unread emails
* Sent emails
* Deleted emails

### Linked Services

Service detection is based on configurable sender matching inside inbox data.

---

## Error Handling

Handles:

* Invalid credentials
* OAuth failures
* Request timeouts
* JSON parsing issues
* Missing fields
* Connection errors

---

## Security Notes

* Intended for educational and authorized testing only
* Do not use against accounts you do not own or have permission to test
* Respect Microsoft Terms of Service
* Use responsibly and legally

---

## Future Improvements

* Async support
* Proxy rotation
* Better logging
* CLI arguments
* Token refresh support
* Multi-threading
* Export formats (CSV/SQLite)

---

## Disclaimer

This project is provided for educational and research purposes only. The author is not responsible for misuse, unauthorized access, or violations of applicable laws or platform policies.

---

## License

MIT License

```txt
MIT License © 2026 DudeGeorgesTG
```
