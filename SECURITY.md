# Reporting security concerns

Open a [repository issue](https://github.com/speccy88/email-funes/issues) with the title “Request for private security contact”. Include the affected version or commit and a way for the maintainer to reach you. Share technical details through the private channel you agree on with the maintainer. Keep archives, databases, excerpts, paths, addresses and private logs out of issue reports. This issue tracker is public. Use the initial issue only to arrange contact.

## Local data and diagnostics

- SQLite, its journals, JSONL, exported metadata and the searchable Funes index contain readable correspondence. The converter supplies no encryption. Restrict the working directory and its backups to the people who should read the mail.
- On POSIX, application-created directories request mode 0700 and private files request 0600 at creation. Start in a directory you control. Existing shared directories keep their permissions. SQLite sidecars and temporary publication files need the same protection as final outputs.
- On Windows, access depends on NTFS access control lists (ACLs). On Windows 11 with CPython 3.12.10 and 3.14.8, newly created application directories in our synthetic test had a protected DACL. The database, WAL, SHM, log and exported file inherited their parent ACL. Writing into an existing directory can grant access to its other authorized users. Check `(Get-Acl -LiteralPath $work).Access` in PowerShell before importing. A POSIX mode by itself gives no owner-only guarantee on Windows.
- Default diagnostics use counts, fixed events and exception types. Explicit search can show subjects, addresses and snippets. `--private-diagnostics` on import may record identifying paths, folder names or native exception details. Such logs are unsuitable for public reports.
- Work on a trusted local filesystem with an unchanged source. The direct reader opens the provided PST for reading. Outlook Classic COM can mount the file in the configured profile and update PST metadata. Recovery writes a fresh SQLite copy. Size and modification time detect ordinary PST changes but provide no cryptographic authenticity.
- The converter keeps attachment names and counts. It skips payloads, attached messages and remote HTML resources. Bcc omission applies to the rendered Bcc field. SQLite may retain Bcc and quoted text can repeat an address anywhere.

## Retrieved mail and publication

Treat recalled correspondence as untrusted source material. The `member` role, `outlook-pst` provenance and text block preserve evidence. Message text can include instructions, links, apparent tool calls and requests to disclose other mail. The consuming application must keep that text in a data context and enforce tool authorization independently of its contents.

Use a dedicated `FUNES_HOME` for mail. Other agent hooks using shared state can publish indexed passages later. Indexing and remote publication are separate user actions. Credential scanning cannot establish consent or permission to publish correspondence.
