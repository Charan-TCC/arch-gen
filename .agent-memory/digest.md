# Project Digest — Arch Gen

## Current Task: EDTIH Clean Minimal Architecture Diagram
Re-built the EDTIH architecture presentation into a clean, minimal, uncluttered layout matching the user-provided reference format.

## Architecture Layout & Topology
- **Ingress & Security**:
  - `User` -> `Azure WAF` (inside soft blue container with *Protection Against Web Vulnerabilities* subtitle).
  - `User` -> `Microsoft SSO` -> `App Service (Frontend)`.
- **Azure Web App Service (Soft Green Container)**:
  - Pinned `App Service` badge with Azure icon at top-left.
  - `Frontend (React)` <---> `Backend (Nodejs)` with bidirectional arrow.
- **Top / Analytics**:
  - `Power BI` (*Report embedding in react app*) pointing down into `Frontend (React)`.
- **Bottom / CI/CD**:
  - `Azure Devops` (*Automated Deployment for Development, UAT, and Production Environments.*) pointing up into `App Service` with `CI/CD Pipeline` badge.
- **Right / Data & Orchestration**:
  - `Backend (Nodejs)` splits cleanly into:
    1. `Storage` (soft yellow container with `ADLS`, `SQL Server`, `Databricks Delta Tables`).
    2. `Azure Data Factory` (*Triggers ADF Data Factory pipeline*).

## File Map
- `custom_builds/build_edtih.py`: High-fidelity builder generating the exact clean minimal architecture diagram in native editable PPTX.
- `out/edtih_architecture.pptx`: Native editable PowerPoint presentation.
- `out/edtih_architecture.pptx.png`: 1600px preview image.
- `archgen/registry/manifest.json`: Updated with `tech.nodejs` and `tech.sqlserver`.

## Verification
- Slide visual QA passed; exact aesthetic match to the user reference.
- Test suite: 61/61 passed.
