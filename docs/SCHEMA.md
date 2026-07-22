# Google Sheet schema

The application stores operational records only in the customer-controlled Google Sheet.

| Sheet | Purpose | Key columns |
| --- | --- | --- |
| `Settings` | Company identity and deployment configuration | Key, Value |
| `Clients` | Delivery recipients | Client ID, Name, Company, Phone, Address, Created At |
| `PackingSlips` | Active slips and archived revisions | Slip Code, Date, Client fields, Items JSON, Quantity, Notes, Invoice Numbers, audit fields, Linked Orders |
| `Users` | Application role allowlist | User ID, Name, Email, Role, Created At |
| `Orders` | Customer orders and delivery status | Order ID, Date, Client fields, Items JSON, Status, Created At, Invoice Numbers |

Do not reorder or rename columns manually. Make a backup before changing schema or importing data. Values that begin with spreadsheet formula markers are persisted as literal text.
