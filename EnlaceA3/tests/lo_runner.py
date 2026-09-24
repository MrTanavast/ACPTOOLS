import sys, time, subprocess, uno, os
from com.sun.star.beans import PropertyValue
def pv(n,v):
    p=PropertyValue(); p.Name=n; p.Value=v; return p
def start():
    """Arranca un LibreOffice propio (puerto y perfil únicos: se pueden lanzar varias pruebas a la vez)."""
    import socket, tempfile
    sock = socket.socket(); sock.bind(("localhost", 0)); port = sock.getsockname()[1]; sock.close()
    perfil = tempfile.mkdtemp(prefix="lo_perfil_")
    proc=subprocess.Popen(["soffice","--headless","--invisible","--norestore","--nologo",
        f"-env:UserInstallation=file://{perfil}",
        f'--accept=socket,host=localhost,port={port};urp;'],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
    proc.perfil = perfil
    local=uno.getComponentContext()
    res=local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver",local)
    for i in range(120):
        try:
            ctx=res.resolve(f"uno:socket,host=localhost,port={port};urp;StarOffice.ComponentContext"); return proc,ctx
        except Exception: time.sleep(0.5)
    proc.terminate()
    raise SystemExit("no arranca soffice")
def run(modules, entry, args=()):
    proc,ctx=start()
    try:
        smgr=ctx.ServiceManager
        desk=smgr.createInstanceWithContext("com.sun.star.frame.Desktop",ctx)
        doc=desk.loadComponentFromURL("private:factory/scalc","_blank",0,(pv("Hidden",True),))
        bl=doc.BasicLibraries
        if not bl.hasByName("Standard"): bl.createLibrary("Standard")
        lib=bl.getByName("Standard")
        for name,code in modules:
            lib.insertByName(name,code)
        sp=doc.getScriptProvider()
        s=sp.getScript(f"vnd.sun.star.script:Standard.{entry}?language=Basic&location=document")
        r=s.invoke(tuple(args),(),())
        return r[0]
    finally:
        try: doc.close(True)
        except Exception: pass
        proc.terminate()
        try:
            proc.wait(timeout=20)
        except Exception:
            proc.kill()
        import shutil
        shutil.rmtree(proc.perfil, ignore_errors=True)
if __name__=="__main__":
    mods=[]
    for f in sys.argv[2:]:
        mods.append((os.path.splitext(os.path.basename(f))[0], open(f,encoding='utf-8').read()))
    print(run(mods, sys.argv[1]))
