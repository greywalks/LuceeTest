# Logicore Portal

Logicore Portal is an internal operations application for invoice generation, nonconforming-material tracking, employee training records, and inventory lifecycle reporting. This repository contains the Lucee/CFML implementation of the portal and preserves the established browser workflows and API contracts from the original Flask application.

## Project status

| Area | Status |
| --- | --- |
| Lucee/CFML conversion | Complete |
| Functional parity validation | Complete |
| Automated Lucee smoke tests | Passing |
| Authenticated workflow tests | Passing |
| Production security configuration | Environment-specific; see **Production checklist** |

Detailed conversion coverage is recorded in [`MIGRATION_PARITY.md`](MIGRATION_PARITY.md).

## Core capabilities

### Invoice Generator

- Promethean Workshop invoices
- Promethean Storage and Small Parts invoices
- Promethean FedEx shipment uploads
- AMC warehouse invoices
- TCL warehouse invoices
- Philips warehouse and repair invoices
- Pricing, dimensions, and reference-data administration

### Operations modules

- **SMS NonConforming:** record management, search, numbering, XLSX exports, and Zebra ZPL labels
- **Training Tracker:** training weeks, topics, videos, sessions, rosters, attendance, signatures, scanned sign-off sheets, and reports
- **Inventory Management:** receiving, shipping, repair, and FedEx imports; file and event deduplication; lifecycle history; shipping reports; quality audits; overrides; whitelist management; and CSV/XLSX exports
- **Administration:** portal users, section-level permissions, Training Tracker roles, and generated-file access controls

## Technology

- Lucee 7 and CFML
- CommandBox for local runtime management
- H2 for persistent application data
- Apache POI through Lucee/Java interoperability for Excel processing
- Tailwind CSS, CSS, and vanilla JavaScript for the existing frontend

## Local development

### Prerequisites

- [CommandBox](https://www.ortussolutions.com/products/commandbox)
- A supported Java runtime for the installed CommandBox/Lucee version
- Python 3.12 only when running the automated parity client

From the repository root:

```bash
box install commandbox-cfconfig --force
box server start
```

The local server opens at:

```text
http://127.0.0.1:5000
```

The runtime is configured in [`server.json`](server.json).

### First-run account

An empty database creates the following bootstrap administrator:

```text
Username: admin
Password: admin
```

> **Security requirement:** This account is intended only for initial setup. Change its password immediately after the first sign-in. Do not expose an instance using the bootstrap credentials to any shared or public network.

## Runtime data

The application creates and updates the following local directories:

| Directory | Purpose |
| --- | --- |
| `data/` | H2 database and uploaded training sign-off records |
| `uploads/` | Temporary source files received by portal workflows |
| `outputs/` | Generated invoices, reports, and access metadata |
| `config/` | Versioned defaults plus runtime reference-data overrides |

Runtime databases, uploads, generated outputs, and mutable configuration files are excluded from version control. Back up the persistent directories according to the deployment's retention requirements.

## Validation

The automated validation suite starts a real Lucee instance and verifies:

- Health, login, portal navigation, and permission enforcement
- All invoice analysis, confirmation, generation, and download workflows
- Workbook sheet names, headers, formulas, metadata cells, dates, and identifier types
- Workshop billing history, duplicate exclusion, price tiers, and companion workbooks
- NonConforming CRUD, exports, numbering, and label output
- Training attendance, PDFs, signed-sheet uploads, ZIP/XLSX reports, and roles
- Inventory imports, duplicate detection, lifecycle views, reporting, and restricted-user access

The test definitions are located in:

- [`.github/workflows/lucee-smoke-test.yml`](.github/workflows/lucee-smoke-test.yml)
- [`tests/billing_parity.cfm`](tests/billing_parity.cfm)
- [`tests/http_parity.py`](tests/http_parity.py)

The `.github/workflows` files are retained for the upstream GitHub validation workflow. GitLab does not execute GitHub Actions; configure a GitLab CI pipeline separately if validation must run natively in this project.

## Production checklist

Before deploying beyond a developer workstation:

- Replace the bootstrap administrator password and create named administrator accounts.
- Serve the portal through HTTPS and enable secure session cookies in `Application.cfc`.
- Keep `data/`, `uploads/`, `outputs/`, application source, and configuration files inaccessible from direct web requests.
- Provide persistent storage and backups for the H2 database and required uploaded records.
- Limit filesystem permissions to the service account running Lucee.
- Set upload, request-size, reverse-proxy, and retention limits appropriate for production workloads.
- Run the full authenticated parity suite against an isolated test instance after deployment changes.

## Repository structure

| Path | Responsibility |
| --- | --- |
| `Application.cfc` | Application bootstrap, datasource, sessions, service registration, and request protection |
| `index.cfm` | Main router and API surface |
| `routes/` | Training Tracker and Inventory Management route handlers |
| `services/` | Authentication, configuration, invoices, Excel, reporting, training, inventory, and output services |
| `static/` | Frontend CSS, JavaScript, and images |
| `views/` | Lucee-compatible portal shell |
| `template/` | Invoice templates and source-compatible workbook layouts |
| `config/` | Reference data and default configuration |
| `lib/` | Pinned Apache POI runtime dependencies |
| `migration/` | Conversion-time templates and rendering utilities |
| `tests/` | Billing fixtures and authenticated end-to-end parity tests |

## Compatibility and source synchronization

The Lucee implementation retains the existing route and browser contracts, including invoice, NonConforming, Training Tracker, and Inventory Management endpoints. This allows the established frontend to operate without a simultaneous UI rewrite.

The upstream asset-import workflow can synchronize unchanged frontend assets, workbook templates, logos, reference JSON, and rendered portal markup from the original Flask repository. Python/Jinja is used only by that conversion-time workflow; the deployed application itself runs on Lucee/CFML.
