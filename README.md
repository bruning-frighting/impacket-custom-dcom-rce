# impacket-custom-dcom-rce

Custom Impacket-based DCOM RCE covering `MMC20.Application`, `Excel.Application`,
`Outlook.Application`, `ShellWindows`, and `ShellBrowserWindow` — raw DCE/RPC
`IDispatch` calls (no Windows COM API / .NET interop), same pattern
`dcomexec.py` uses for MMC20/ShellWindows, extended to Excel/Outlook and fixed
for ShellWindows/ShellBrowserWindow.

> ⚠️ **Authorized use only.** This is a lateral-movement / remote-command-execution
> tool intended for use during penetration tests, red-team engagements, or lab
> research on systems you own or are explicitly authorized to test. Do not run
> this against systems without written authorization.

## What it does

Output is captured the same way `dcomexec.py` does: the remote command is
wrapped to redirect stdout/stderr to the target's own `ADMIN$` share via the
`127.0.0.1` loopback, then the host reads that file back over its own SMB
session to the target and deletes it. Output prints straight to your host
console — no manual file checking on the target.

### Preconditions

- `shellwindows` / `shellbrowserwindow` require at least one Explorer folder
  window already open in the target's interactive session. Without it,
  `.Item()` / the ROT lookup returns nothing and the chain aborts.
- `ie` (`InternetExplorer.Application`) activation is **not** reproducible
  remotely on Vista+/Server 2008+ targets by design (Session 0 isolation —
  RPCSS cannot create a new process running as the interactive user for a
  network activation request). The `ie` technique here is a DLL search-order
  hijack instead: it plants a DLL at `C:\Program Files\Internet Explorer\iertutil.dll`
  and triggers a load via DCOM activation, not a shell-command primitive. See
  the module docstring in `impacket_custom_rce.py` for the full root-cause
  writeup.

## Install

```bash
git clone https://github.com/<your-username>/impacket-custom-dcom-rce.git
cd impacket-custom-dcom-rce
python3 -m pip install -r requirements.txt
```

Requires Python 3 and network access (SMB/135/RPC) to the target.

## Usage

```
python3 impacket_custom_rce.py <mmc20|excel|excel_xll|outlook|outlook_scriptcontrol|shellwindows|shellbrowserwindow|ie> \
    <target_ip> <username> <password> [domain] [command-or-dll-path-or-jscript]
```

- `<command>` is the OS command to run, e.g. `whoami` or `ipconfig /all`.
  Defaults to `whoami` if omitted.
- For `excel_xll` / `ie`, the 6th argument is a **local file path** (the
  `.xll` to `RegisterXLL`, or the `iertutil.dll` to plant via DLL
  search-order hijack), **not** a command. This repo does not ship those
  payload files — bring your own.
- For `ie`, pass `-` or `skip` as the 6th argument to skip the SMB upload
  and reuse a DLL already planted on the target from a previous run.
- For `outlook_scriptcontrol`, the 6th argument is raw JScript source
  (executed synchronously via `ScriptControl.AddCode()`), not a shell
  command. Defaults to a JScript snippet that spawns `WScript.Shell` to
  write a marker file.

### Examples

```bash
python3 impacket_custom_rce.py mmc20 10.0.0.5 administrator 'P@ssw0rd' '' "whoami"
python3 impacket_custom_rce.py excel 10.0.0.5 administrator 'P@ssw0rd' CORP "ipconfig /all"
python3 impacket_custom_rce.py shellwindows 10.0.0.5 administrator 'P@ssw0rd'
python3 impacket_custom_rce.py excel_xll 10.0.0.5 administrator 'P@ssw0rd' '' /path/to/payload.xll
python3 impacket_custom_rce.py ie 10.0.0.5 administrator 'P@ssw0rd' '' /path/to/iertutil.dll
```

Set `RCE_DEBUG=1` in the environment for verbose (DEBUG-level) logging.

## License

Add a license of your choice (e.g. MIT/BSD) before making this repository
public.
