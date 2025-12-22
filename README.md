# MultiDocTemplate

Group-based document templates and archive generation module for Dolibarr 23.0.0-alpha.

## Features

- **User Group Templates**: Upload document templates (ODT, ODS, XLS, XLSX, DOC, DOCX, PDF, RTF) organized by user group
- **Archive Generation**: Generate documents from templates with Dolibarr's variable substitution engine
- **Tag-based Organization**: Archives are organized in folders based on thirdparty/contact categories
- **Native Integration**: Tabs appear directly on User Group, Third Party, and Contact cards
- **Upgrade-safe**: Uses only hooks and $this->tabs - no core file modifications

## Requirements

- Dolibarr 23.0.0-alpha or later
- PHP 7.4+
- Required modules: Societe (Third Parties), User

## Installation

1. Extract the `multidoctemplate` folder to `htdocs/custom/` in your Dolibarr installation
2. Go to **Home > Setup > Modules**
3. Find "MultiDocTemplate" under the Tools family
4. Click to enable the module
5. Configure permissions for users/groups

## Directory Structure

```
multidoctemplate/
├── admin/                 # Module configuration pages
├── class/                 # PHP classes
│   ├── template.class.php      # Template management
│   ├── archive.class.php       # Archive management
│   └── documentgenerator.class.php  # Document generation with substitution
├── core/modules/          # Module descriptor
├── langs/                 # Language files
├── sql/                   # Database schema
├── templates.php          # Template upload interface (user groups)
├── archives.php           # Archive generation interface (thirdparty/contacts)
└── index.php              # Module home page
```

## Usage

### Uploading Templates

1. Go to **Users & Groups > Groups**
2. Select a user group
3. Click the **Templates** tab
4. Upload ODT templates with Dolibarr substitution variables

### Template Variables

Use Dolibarr's standard substitution variables in your ODT templates:

**Third Party:**
- `{THIRDPARTY_NAME}` - Company name
- `{THIRDPARTY_ADDRESS}` - Address
- `{THIRDPARTY_ZIP}` - Postal code
- `{THIRDPARTY_TOWN}` - City
- `{THIRDPARTY_COUNTRY}` - Country
- `{THIRDPARTY_EMAIL}` - Email
- `{THIRDPARTY_PHONE}` - Phone
- `{THIRDPARTY_VAT_INTRA}` - VAT number

**Contact:**
- `{CONTACT_FIRSTNAME}` - First name
- `{CONTACT_LASTNAME}` - Last name
- `{CONTACT_FULLNAME}` - Full name
- `{CONTACT_EMAIL}` - Email
- `{CONTACT_PHONE_PRO}` - Professional phone

**Your Company:**
- `{MYCOMPANY_NAME}` - Your company name
- `{MYCOMPANY_ADDRESS}` - Your address
- etc.

**Date/Time:**
- `{DATE_NOW}` - Current date
- `{DATETIME_NOW}` - Current date and time
- `{YEAR}`, `{MONTH}`, `{DAY}` - Date parts

### Generating Archives

1. Go to a Third Party or Contact card
2. Click the **Archives** tab
3. Select a template from the dropdown
4. Optionally select a category filter for folder organization
5. Click **Generate**

## File Storage

- **Templates**: `DOL_DATA_ROOT/multidoctemplate/templates/group_{id}/`
- **Archives**: `DOL_DATA_ROOT/multidoctemplate/archives/{type}_{id}/{tag}/`

## Permissions

| Permission | Description |
|------------|-------------|
| `lire` | Read module |
| `archive_voir` | View archives |
| `archive_creer` | Generate archives |
| `archive_supprimer` | Delete archives |
| `template_voir` | View templates |
| `template_creer` | Upload templates |
| `template_supprimer` | Delete templates |

## Database Tables

- `llx_multidoctemplate_template` - Stores template metadata
- `llx_multidoctemplate_archive` - Stores generated archive metadata

## License

GNU General Public License v3.0

## Support

For issues or feature requests, please contact the developer.
