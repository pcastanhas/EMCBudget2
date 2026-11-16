import os, sys, time, subprocess, tempfile
from pathlib import Path
sys.path.insert(0, "/mnt/skills/public/xlsx/scripts")
sys.path.insert(0, "/usr/lib/python3/dist-packages")
try:  # optional helper (LD_PRELOAD socket shim) when present; plain headless env otherwise
    from office.soffice import get_soffice_env
except ImportError:
    def get_soffice_env():
        env = os.environ.copy()
        env["SAL_USE_VCLPLUGIN"] = "svp"
        return env
import uno
from com.sun.star.beans import PropertyValue

PORT = 2002


def pv(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


class LO:
    def __init__(self):
        self.profile = tempfile.mkdtemp(prefix="lo_prof_")
        args = ["soffice", f"-env:UserInstallation={Path(self.profile).as_uri()}", "--headless", "--invisible",
                "--norestore", "--nologo", f"--accept=socket,host=localhost,port={PORT};urp;"]
        self.proc = subprocess.Popen(args, env=get_soffice_env(), stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
        last = None
        for _ in range(60):
            try:
                self.ctx = resolver.resolve(f"uno:socket,host=localhost,port={PORT};urp;StarOffice.ComponentContext")
                break
            except Exception as e:
                last = e
                time.sleep(1)
        else:
            raise RuntimeError(f"cannot connect: {last}")
        self.smgr = self.ctx.ServiceManager
        self.desktop = self.smgr.createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)

    def load(self, path):
        url = uno.systemPathToFileUrl(os.path.abspath(path))
        return self.desktop.loadComponentFromURL(url, "_blank", 0, (pv("Hidden", True),))

    def add_module(self, doc, mod_name, source):
        libs = doc.BasicLibraries
        if not libs.hasByName("Standard"):
            libs.createLibrary("Standard")
        lib = libs.getByName("Standard")
        libs.loadLibrary("Standard")
        if lib.hasByName(mod_name):
            lib.removeByName(mod_name)
        lib.insertByName(mod_name, source)

    def run_macro(self, doc, macro):
        """macro like 'Standard.BuildMod.BuildAll' (library.module.sub)"""
        sp = doc.getScriptProvider()
        script = sp.getScript(f"vnd.sun.star.script:{macro}?language=Basic&location=document")
        return script.invoke((), (), ())

    def save(self, doc, path, filt):
        url = uno.systemPathToFileUrl(os.path.abspath(path))
        doc.storeToURL(url, (pv("FilterName", filt), pv("Overwrite", True)))

    def close(self):
        try:
            self.desktop.terminate()
        except Exception:
            pass
        time.sleep(1)
        try:
            self.proc.kill()
        except Exception:
            pass
