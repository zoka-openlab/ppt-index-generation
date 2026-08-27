# Security and Privacy

## Supported versions

Security and privacy fixes are applied to the latest version on the default branch.

## Reporting a vulnerability

Please report vulnerabilities privately to the repository maintainer instead of opening a public issue. Include the affected version, reproduction steps, and potential impact. Do not attach real or confidential presentations.

## Data handling

The Skill runs locally and reads only paths supplied by the user. It does not upload presentations by itself. Generated Markdown indexes may contain:

- absolute file paths
- slide titles and extracted text
- file size and modification time
- slide-level structure metadata

Treat generated indexes and candidate packs as sensitive when the source material is sensitive. Store them in an access-controlled local folder and review them before sharing.
