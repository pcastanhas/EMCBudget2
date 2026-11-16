import os, uno
from lo_driver import LO, pv
from com.sun.star.table import CellRangeAddress
lo=LO()
try:
    import sys; doc=lo.load(sys.argv[1])
    areas={"COVER":(0,0,11,62),"INPUT":(0,0,4,44),"DETAIL":(0,0,20,22),"QTRLY SUM":(0,0,8,24),"REVENUE":(0,0,20,50),"OTHER_REVENUE":(0,0,20,50)}
    import os
    if os.environ.get("RENDER_OPT"):
        doc.Sheets.getByName("INPUT").getCellRangeByName("C31").setString(os.environ["RENDER_OPT"]); doc.calculateAll()
    ps=doc.StyleFamilies.getByName("PageStyles")
    for i in range(doc.Sheets.Count):
        sh=doc.Sheets.getByIndex(i); nm=sh.Name
        a=CellRangeAddress(); a.Sheet=i
        if nm in areas:
            c1,r1,c2,r2=areas[nm]; a.StartColumn,a.StartRow,a.EndColumn,a.EndRow=c1,r1,c2,r2
        else:
            a.StartColumn=a.StartRow=a.EndColumn=a.EndRow=0
        sh.setPrintAreas((a,))
        st=ps.getByName(sh.PageStyle); st.IsLandscape=True
        w,h=st.Width,st.Height
        if w<h: st.Width,st.Height=h,w
        st.ScaleToPagesX=1; st.ScaleToPagesY=1
    url=uno.systemPathToFileUrl(os.path.abspath("render_cover.pdf"))
    doc.storeToURL(url,(pv("FilterName","calc_pdf_Export"),))
    doc.close(True)
finally:
    lo.close()
