# MultiDocTemplate - Dolibarr Module

Group-based document templates and archive generation module for Dolibarr 23.0+

## Features

- Upload document templates by user group
- Generate documents from Third Party and Contact cards
- Variable substitution in ODT, DOCX, XLSX, ODS formats
- Tag-based folder organization
- File explorer style interface with collapsible folders
- Search and sort functionality

## Supported Formats

| Format | Variable Substitution | Notes |
|--------|----------------------|-------|
| ODT | Yes | LibreOffice Writer - Full support |
| ODS | Yes | LibreOffice Calc - Full support |
| DOCX | Yes | Microsoft Word - Full support |
| XLSX | Yes | Microsoft Excel - Full support |
| PDF | No | Copy only |
| DOC | No | Copy only |
| XLS | No | Copy only |
| RTF | No | Copy only |

## Substitution Variables

### Thirdparty (Company) Variables
| Variable | Description |
|----------|-------------|
| {company_name} | Company name |
| {company_name_alias} | Company alias |
| {company_address} | Address |
| {company_zip} | Postal code |
| {company_town} | City |
| {company_country} | Country name |
| {company_country_code} | Country code (ES, FR, US...) |
| {company_state} | State/Province |
| {company_phone} | Phone |
| {company_fax} | Fax |
| {company_email} | Email |
| {company_web} | Website |
| {company_customercode} | Customer code |
| {company_idprof1} to {company_idprof6} | Professional IDs |
| {company_vatnumber} | VAT number |
| {company_capital} | Capital |
| {company_note_public} | Public note |
| {company_note_private} | Private note |
| {company_default_bank_iban} | Bank IBAN |
| {company_default_bank_bic} | Bank BIC |
| {company_logo} | Company logo URL |
| {company_options_XXX} | Extra fields (replace XXX with field code) |

### Contact Variables
| Variable | Description |
|----------|-------------|
| {contact_civility} | Title (Mr, Mrs...) |
| {contact_firstname} | First name |
| {contact_lastname} | Last name |
| {contact_fullname} | Full name |
| {contact_poste} | Job position |
| {contact_address} | Address |
| {contact_zip} | Postal code |
| {contact_town} | City |
| {contact_phone} | Phone |
| {contact_phone_mobile} | Mobile |
| {contact_email} | Email |
| {contact_birthday} | Birthday |
| {contact_photo} | Contact photo URL |
| {contact_options_XXX} | Extra fields |

### Logged-in User Variables
| Variable | Description |
|----------|-------------|
| {user_login} | Username |
| {user_firstname} | First name |
| {user_lastname} | Last name |
| {user_fullname} | Full name |
| {user_email} | Email |
| {user_phone} | Office phone |
| {user_phone_mobile} | Mobile |
| {user_signature} | Signature |
| {user_job} | Job title |
| {user_options_XXX} | Extra fields |

### My Company Variables
| Variable | Description |
|----------|-------------|
| {mycompany_name} | Company name |
| {mycompany_address} | Address |
| {mycompany_zip} | Postal code |
| {mycompany_town} | City |
| {mycompany_country} | Country |
| {mycompany_phone} | Phone |
| {mycompany_email} | Email |
| {mycompany_vatnumber} | VAT number |
| {mycompany_logo} | Company logo URL |

### Date Variables
| Variable | Description |
|----------|-------------|
| {date} | Current date |
| {datehour} | Current date and time |
| {year} | Year |
| {month} | Month (01-12) |
| {day} | Day (01-31) |
| {current_date} | Current date |
| {current_datehour} | Current date and time |
| {current_date_locale} | Date in locale format |
| {current_datehour_locale} | Date/time in locale format |

## Installation

1. Extract to htdocs/custom/multidoctemplate/
2. Go to Home > Setup > Modules
3. Activate "MultiDocTemplate"

## Usage

### Upload Templates
1. Go to Users & Groups > Groups > select group > Templates tab
2. Fill in Tag (folder) and Label
3. Upload template file (.odt, .docx, .xlsx, etc.)

### Generate Documents
1. Go to Third Parties or Contacts
2. Select a record > Archives tab
3. Select template and click Generate

## Notes

- When generating from Contact, both contact AND company variables are available
- Extra fields: {xxx_options_FIELDCODE}
- Logo variables return URLs for use in HTML templates
- PDF/RTF files are copied without substitution

## License

GPL v3+
