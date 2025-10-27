# OpenProject Timesheet Export

Export OpenProject time entries to beautifully formatted Excel timesheets.

## Features

- Export time entries for any month and user
- Professional Excel formatting with styled headers and borders
- Support for custom fields (e.g., location tracking)
- Automatic pagination for large datasets
- ISO 8601 duration parsing
- Configurable via CLI arguments or environment variables

## Installation

Using [uv](https://github.com/astral-sh/uv):

```bash
uv pip install git+https://github.com/jeremi/openproject-timesheet-export.git
```

Or install from source:

```bash
git clone https://github.com/jeremi/openproject-timesheet-export.git
cd openproject-timesheet-export
uv pip install -e .
```

## Usage

### Basic Usage

```bash
openproject-export --base-url https://openproject.example.com --api-key YOUR_API_KEY
```

### With Environment Variables

Set up your environment:

```bash
export OPENPROJECT_BASE_URL=https://openproject.example.com
export OPENPROJECT_API_KEY=YOUR_API_KEY
export OPENPROJECT_USER=me
```

Then simply run:

```bash
openproject-export
```

### Advanced Options

Export a specific month:

```bash
openproject-export --month 2024-01
```

Export for a specific user:

```bash
openproject-export --user 42
```

Specify custom field for location tracking:

```bash
openproject-export --location-cf customField7
```

Custom output path:

```bash
openproject-export --out /path/to/timesheet.xlsx
```

## Command Line Options

| Option | Environment Variable | Default | Description |
|--------|---------------------|---------|-------------|
| `--base-url` | `OPENPROJECT_BASE_URL` | - | OpenProject instance URL (required) |
| `--api-key` | `OPENPROJECT_API_KEY` | - | API key for authentication (required) |
| `--month`, `-m` | - | Current month | Month to export (format: YYYY-MM) |
| `--user` | `OPENPROJECT_USER` | `me` | User ID or 'me' for current user |
| `--location-cf` | `OPENPROJECT_LOCATION_CF` | - | Custom field key for location (e.g., customField7) |
| `--out` | - | `timesheet-YYYY-MM.xlsx` | Output file path |
| `--page-size` | - | 200 | API page size for pagination |

## Output Format

The exported Excel file includes:

- **Date**: The date the work was performed
- **Working hours**: Duration in decimal hours
- **Location**: Location of work (from custom field or default "remote")
- **Assignment number_Activity_Work content**: Combined field with assignment number, activity name, and comments

### Excel Formatting

- Professional blue header with white text
- Center-aligned date, hours, and location columns
- Left-aligned, wrapped text for assignment details
- Borders on all cells
- Optimized column widths

## Getting Your API Key

1. Log in to your OpenProject instance
2. Go to **My account** (top-right menu)
3. Select **Access tokens** from the left menu
4. Click **+ API token** to create a new token
5. Copy the generated token

## Requirements

- Python 3.11 or higher
- OpenProject instance with API access

## License

MIT License - see LICENSE file for details

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Support

If you encounter any issues, please [open an issue](https://github.com/jeremi/openproject-timesheet-export/issues) on GitHub.
