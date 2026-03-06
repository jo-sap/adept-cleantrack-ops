# CleanTrack Ops

CleanTrack Ops is an internal operations management platform developed for **Adept Services Pty Ltd**.
The system provides a centralised interface for managing operational data including cleaning sites, staff, timesheets, compliance records, and internal workflows.

The platform integrates with **Microsoft 365 and SharePoint Lists**, which act as the primary data source, allowing Adept's team to manage operational records in a structured and scalable way.

## Key Features

* Site management across multiple states
* Cleaner and staff management
* Timesheet and attendance tracking
* Compliance and operational oversight
* Integration with Microsoft authentication
* SharePoint Lists used as the operational database
* Real-time data updates across the organisation

## Architecture

CleanTrack Ops follows a lightweight cloud architecture:

Frontend Application
Cloudflare Pages hosting
Microsoft Authentication (Azure AD)
Microsoft Graph API
SharePoint Lists (Data Storage)

This architecture allows the platform to remain cost-efficient while leveraging the security and reliability of the Microsoft ecosystem.

## Deployment

The application is deployed via **Cloudflare Pages** and connected to this GitHub repository.

Development workflow:

* `dev` branch – development and testing
* `main` branch – production deployment

Updates pushed to GitHub automatically trigger deployments through Cloudflare Pages.

## Purpose

CleanTrack Ops was built to improve operational visibility, reduce administrative workload, and provide Adept Services with a scalable internal platform to support company growth and multi-site service management.

## Author

Developed internally for **Adept Services Pty Ltd**.
