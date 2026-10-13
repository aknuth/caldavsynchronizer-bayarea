# Updates and releases

How a new version gets from this repository to the customers.

## How the add-in finds a new version

There is no automatic update. The user clicks **About → Check for updates**.

1. The add-in downloads
   `https://download.bayarea-cc.com/caldavsynchronizer/latest.json`:
   ```json
   {
     "version": "5.1.0",
     "url": "https://download.bayarea-cc.com/caldavsynchronizer/CalDavSynchronizer-BayArea-5.1.0.zip",
     "notes": "Short text shown to the user."
   }
   ```
2. It compares `version` with its own assembly version.
3. If the version on the server is newer, it shows the version and `notes` and
   offers to open `url` in the browser. Otherwise it reports *No newer version
   available*.
4. The user unzips the download and runs `setup.exe`. The installer replaces
   the installed version (MSI major upgrade); the profiles are kept.

The address of `latest.json` is `WebResourceUrls.SiteContainingNewestVersion`.
The background update check of the original add-in is switched off.

## Releasing a new version

| Step | Who | Where |
| --- | --- | --- |
| 1. Raise the version, push | developer | Linux |
| 2. Pull, build, zip | developer | Windows VM |
| 3. Upload zip, then `latest.json` | developer | Windows VM |
| 4. Check for updates in Outlook | developer | Windows VM |
| 5. Tag the release commit, push the tag | developer | Linux |

### 1. Raise the version

All in one commit `Version X.Y.Z`:

- `CalDavSynchronizer.Setup/CalDavSynchronizer.Setup.vdproj`:
  `ProductVersion`, plus a **new** `ProductCode` and `PackageCode`. The
  `UpgradeCode` stays the same, otherwise the old version isn't replaced.
- `CalDavSynchronizer/Properties/AssemblyInfoVersion.cs`: `AssemblyVersion`.
  This is the version the update check compares.
- `release/latest.json`: `version`, `url` and `notes`.

### 2. Build

PowerShell in the working copy, one command per line (Windows PowerShell has
no `&&`):

```powershell
git pull
```

In Visual Studio: configuration **Release**, *Rebuild Solution*, then build
`CalDavSynchronizer.Setup`. The setup takes its files from `bin\Release`, so a
Debug build gives an installer with old or mismatched files.

```powershell
Compress-Archive -Path CalDavSynchronizer.Setup\Release\setup.exe, CalDavSynchronizer.Setup\Release\CalDavSynchronizer.Setup.msi -DestinationPath CalDavSynchronizer-BayArea-X.Y.Z.zip
```

Both files are needed: `setup.exe` only installs the prerequisites and runs
the `.msi` next to it.

### 3. Upload

```powershell
scp CalDavSynchronizer-BayArea-X.Y.Z.zip <user>@<download server>:<download directory>/caldavsynchronizer/
scp release\latest.json <user>@<download server>:<download directory>/caldavsynchronizer/
```

Zip first, then `latest.json`, so `latest.json` never points to a missing
file. The version in the file name keeps browsers and proxies from serving an
old cached zip. Old zips can stay on the server.

### 4. Check

In an Outlook with the previous version: *About → Check for updates* shows
the new version and the notes.

### 5. Tag

On the commit the release was built from:

```bash
git tag -a vX.Y.Z <commit> -m "CalDavSynchronizer (Bay Area Edition) X.Y.Z"
git push bayarea vX.Y.Z
```

The AGPL requires the source of every released version to be available; the
tag is how to find it.
