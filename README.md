# expd — Expedition Data System

Web-based pre- and post-cruise data entry and query system for MBARI research expeditions.

## Overview

`expd` is a classic ASP / VBScript application backed by a SQL Server database. It provides:

- **Pre-cruise forms** — dive planning, waypoints, personnel registration
- **Post-cruise forms** — dive logs, CTD/ROV data review, annotations, camera logs
- **Nightly download scripts** — automated sync from shipboard systems (VARS, SSDS, etc.)
- **Query interfaces** — expedition data retrieval for scientists and data managers

## Stack

| Component  | Technology             |
|------------|------------------------|
| Web server | IIS (Windows Server)   |
| Language   | ASP Classic / VBScript |
| Database   | SQL Server             |
| Client     | HTML + JavaScript      |

## Directory structure

```
expd/
├── log/           Dive log and cruise forms (pre/postcruise, waypoints, etc.)
├── queries/       Data query pages (annotations, nav, CTD, camera logs, KML)
├── styles/        CSS stylesheets
├── docs/          Documentation and design diagrams
├── UserDocs/      End-user documentation (HTML)
├── bin/           Server-side DLLs (RichDatePicker, msgBox)
└── Web.config     IIS configuration
```

## Configuration

Application settings (database connection, logistics email, registered users) are in:

```
log/expd_functions.inc
```

## Status

**Maintenance-only.** Not under active development.

## Source of truth and updates

The running application lives on the production **wwwroot** server. This GitHub repository is the institute archive of that codebase.

Edits are made on the wwwroot server. Changed files are then pushed from the server to this repository (via rsh). GitHub is not the editing environment.

Write access to this repository is limited to the archive maintainer.

Database credentials in this GitHub copy are masked. Live connection strings remain on the server and should stay there when files are pushed up.

## Full archive

Complete history including dev branch and all backups:
[ksalamy/expd-archive](https://github.com/ksalamy/expd-archive)

## Contact

MBARI Data Team — [www.mbari.org](https://www.mbari.org)
