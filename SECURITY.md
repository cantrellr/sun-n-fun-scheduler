# Security Policy

## Scope

This repository includes operational schedule data and volunteer contact details.
Security issues in this project can affect confidentiality, integrity, or safe
handling of that information.

## Reporting A Security Issue

Do not open a public issue with:

- personal contact information
- raw data extracts
- screenshots containing volunteer details
- exploit details that increase immediate risk

Instead, report the issue privately through the repository owner, maintainer, or
your approved internal support channel.

## What To Include In A Report

- affected file or workflow
- short description of the issue
- steps to reproduce
- expected behavior
- actual behavior
- any mitigation already identified

## Data Handling Expectations

- Treat `data/` files as sensitive working content.
- Treat generated exports as sensitive operational artifacts.
- Sanitize logs and screenshots before sharing.
- Avoid sending schedule data over unsecured or public channels.

## Current Security Limitations

- no authentication model
- no authorization model
- no audit log
- no encrypted storage model inside the repository
- no automatic secret scanning or CI enforcement configured in this repo

## Recommended Operational Safeguards

- restrict repository access to approved users
- avoid broad redistribution of exported files
- keep working copies on managed devices where possible
- back up the working schedule file before major edits
- move toward a service-backed architecture before expanding multi-user use
