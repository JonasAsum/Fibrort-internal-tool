# CRM System

Tools that get sales data into the Margince CRM.

```
importer/       one site: Excel lists (visits, customers, contacts, to-dos, orders, receivables) and Thunderbird mail -> the CRM (see importer/README.md)
mock_visits/    Invented visit workbooks for tests and demos, no real data
imap-testbed/   Local IMAP server for trying the mail importer
```

`margince/` (the CRM itself) is a separate project and is not part of this repo.
Real customer data (`Email Data/`, `Excel Files from Sales Meetings/`, `importer/mail/staging/`) and credentials are git-ignored.

Setup and commands: `importer/README.md`.
