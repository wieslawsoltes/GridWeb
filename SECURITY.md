# Security

Treat spreadsheets and imported ZIP/XML as untrusted data. GridWeb parses formulas with a bounded interpreter rather than eval/Function; unknown functions yield explicit errors. ZIP readers limit entries and expanded bytes, verify CRCs, reject encryption and unsafe paths. XML input does not load external entities. Host RPC routes only named operations and never arbitrary object paths or scripts.

Native hosts load an embedded local page, deny unrelated navigation and new windows, and keep operating-system file/network capabilities outside workbook formulas. Do not replace the host page with remote content while exposing its native transport.

The baseline does not execute VBA, DDE, external workbook connections, ActiveX or embedded executables. A workbook is not trusted merely because it has a familiar extension. UI protection is not encryption or access control. JSON/CSV/XLSX exports should be reviewed before distribution; formulas may be evaluated by other spreadsheet applications.

Report suspected vulnerabilities privately to the repository owner. Do not publish secrets or exploit payloads in public issues. Release workflows read NPM_TOKEN/NUGET_API_KEY only in publication steps and verify immutable package integrity.
