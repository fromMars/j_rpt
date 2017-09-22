Param.Value['BARCODE']:=StrRight('0000'+PARAM.Value['_RUNTAG'],4)
+StrRight('00'+PARAM.Value['BATCHREF'],2)
+StrRight('0000'+PARAM.Value['CUTCNT'],4);



/* --debug-- */
nlist:=strings.create();
nlist.add('JOBREF');
nlist.add('RUNTAG');
nlist.add('RUNTAGDEC');
nlist.add('PROJECTREF');
nlist.add('ASSEMBLYREF');
nlist.add('FRAMEREF');
nlist.add('FRAMEOPEN');
nlist.add('VENTREF');
nlist.add('LABS');
nlist.add('LMAX');
nlist.add('LMIN');
nlist.add('LENGTH');
nlist.add('ANGLEB');
nlist.add('ANGLEE');
nlist.add('PROFILE');
nlist.add('SYSTEM');
nlist.add('PRODUCT');
nlist.add('ORDERCODE');
nlist.add('PROFILEDESC');
nlist.add('FINISH');
nlist.add('FINISHDESC');
nlist.add('CARRIER');
nlist.add('CABIN');
nlist.add('LCNT');
nlist.add('PROFILENO');
nlist.add('POSITION');
nlist.add('ENFCODE1');
nlist.add('ENFLABS1');
nlist.add('ENFCODE2');
nlist.add('ENFLABS2');
nlist.add('ENFCODE3');
nlist.add('ENFLABS3');
nlist.add('ENFORCED');
nlist.add('DELIVERYWEEK');
nlist.add('CUSTOMERREF');
nlist.add('CUSTOMERDESC');
nlist.add('GLYPHREF');
nlist.add('POSSYMBOL');
nlist.add('FRAMENO');


fn:=''+getparam('PROGRAM_ROOT')+'\\'+'OUTPUT'+'\\'+getparam('PROJECT_TEXT')+'\\'+'debug3.xml';
xd:=xmldocument.create();
if !xd.loadfile(fn) then
{
    xd.savefile(fn);
    xd:=xmldocument.create();
}

root:=xd.DocumentRoot;
pn:=xmlelement.create();
pn_str:='pn'+getparam('BARCODE');
pn.nodename:=pn_str;
/*pn.value:=1;*/
root.addelement(pn,-1);
ncnt:=nlist.count;
/*pn:=root.getelementbyname(pn_str);*/
while ncnt>0 do
{
    ncnt:=ncnt-1;
    x:=xmlelement.create();
    x.nodename:=nlist.strings[ncnt];
    x.value:=trim(getparam(nlist.strings[ncnt]));
    pn.addelement(x,0);
}

xd.savefile(fn);
