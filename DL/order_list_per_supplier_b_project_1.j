; order_list_per_supplier_b_project_1
; Bestellijst/leverancier: Profielen

%% detail
range := template.Range["RANGE_DETAIL_PROFILE"];
if !IsIDispatch(range) then
{
  errmsg := "Range RANGE_DETAIL_PROFILE not found!";
  goto error;
};
excel.Goto(range);
row := row + 1;
range := range.Offset[row];
excel.Goto(range);
excel.Selection.EntireRow.Insert();
template.Range["RANGE_DETAIL_PROFILE"].Copy();
template.Paste;
range := range.Offset[-1];
if !IsIDispatch(range) then
{
  errmsg := "Range RANGE_DETAIL_PROFILE is incorrect!";
  goto error;
};
cell := range.Cells[1][1];
if IsIDispatch(cell) then cell.Value := @%DB_ATTRIB_NO%;
cell := range.Cells[1][2];
if IsIDispatch(cell) then cell.Value := iif(trim("%DSP_ATTRIB_ARTICLECODE%")="",trim("%DSP_ATTRIB_ACC%")+"."+trim(%IF{%ORDER_LIST_PER_SUPPLIER_VARIANT%=0,"%DSP_ATTRIB_SERIE%","%DSP_ATTRIB_VARIETY%"}),trim("%DSP_ATTRIB_ARTICLECODE%"));
cell := range.Cells[1][3];
if IsIDispatch(cell) then cell.Value := HtmlToNormalStr(trim("%DSP_ATTRIB_VARIETYDESC%"));
cell := range.Cells[1][4];
if IsIDispatch(cell) then cell.Value := HtmlToNormalStr(trim("%DSP_ATTRIB_ACCDESC%"));
if %ORDER_LIST_PER_SUPPLIER_BESTPRICE% then
{
  cell := range.Cells[1][5];
  if IsIDispatch(cell) then cell.Value := @%DB_ATTRIB_LENGTH%;
  cell := range.Cells[1][6];
  if IsIDispatch(cell) then cell.Value := @%DB_ATTRIB_ITMPRICE%;
  cell := range.Cells[1][9];
  if IsIDispatch(cell) then cell.Value := @%DB_ATTRIB_PACKSIZE%;
  cell := range.Cells[1][10];
  if IsIDispatch(cell) then cell.Value := @%DB_ATTRIB_SETPRICE%;
  cell := range.Cells[1][11];
  if IsIDispatch(cell) then cell.Value := iif(%ORDER_LIST_PER_SUPPLIER_REBATE%,(1-@%DB_ATTRIB_REBATE%/100),1);
}
else
{
  cell := range.Cells[1][5];
  if IsIDispatch(cell) then cell.Value := @%DB_ATTRIB_PACKVOLUME%;
  cell := range.Cells[1][6];
  if IsIDispatch(cell) then cell.Value := @%DB_ATTRIB_LENGTH%;
  cell := range.Cells[1][7];
  if IsIDispatch(cell) then cell.Value := @%DB_ATTRIB_PACKCOUNT%;
  cell := range.Cells[1][8];
  if IsIDispatch(cell) then cell.Value := @%DB_ATTRIB_PRICE%;
}

%% break header
; dupliceer sheet "profile" en maak hoofding met volgende info:
tmp_profile.Copy(tmp_profile);
cnt := book.WorkSheets.Count-3;
if cnt <= 0 then
{
  errmsg := "No of sheets in base template is incorrect!";
  goto error;
};
template := book.WorkSheets[cnt];
template.Name := "profile_"+trim("%DSP_ATTRIB_SUPPLIER%");
template.Activate();
cell := template.Range["CELL_PROFILE_SUPPLIER"];
if IsIDispatch(cell) then cell.Value := "%DSP_TEXT_CLIENT%";
cell := template.Range["CELL_PROFILE_SUPPLIERDESC"];
if IsIDispatch(cell) then cell.Value := "%DSP_TEXT_CONTACT%";
cell := template.Range["CELL_PROFILE_SUPPLIERADDRESS1"];
if IsIDispatch(cell) then cell.Value := "%DSP_TEXT_STREET%";
cell := template.Range["CELL_PROFILE_SUPPLIERADDRESS2"];
if IsIDispatch(cell) then cell.Value := trim("%DSP_TEXT_ZIP%")+" "+trim("%DSP_TEXT_PLACE%");
cell := template.Range["CELL_PROFILE_SUPPLIERADDRESS3"];
if IsIDispatch(cell) then cell.Value := "%DSP_TEXT_COUNTRY%";
cell := template.Range["CELL_PROFILE_SUPPLIERPHONE"];
if IsIDispatch(cell) then cell.Value := "%DSP_TEXT_PHONE%";
cell := template.Range["CELL_PROFILE_SUPPLIERFAX"];
if IsIDispatch(cell) then cell.Value := "%DSP_TEXT_TELEFAX%";

; Translation of header price labels.
cell := template.Range["PURCHASE_PRICE_LABEL1"];
if IsIDispatch(cell) then cell.Value := GetLanText(-312) + " (" + "%CURRENCY%" +")";
cell := template.Range["UNIT_PRICE_LABEL1"];
if IsIDispatch(cell) then cell.Value := GetLanText(-26005) + " (" + "%CURRENCY%" +")";
cell := template.Range["GRADUATE_PRICE_LABEL1"];
if IsIDispatch(cell) then cell.Value := GetLanText(-26006) + " (" + "%CURRENCY%" +")";
cell := template.Range["PACKING_PRICE_LABEL1"];
if IsIDispatch(cell) then cell.Value := GetLanText(-26007) + " (" + "%CURRENCY%" +")";

row := 0;
start := template.Range["RANGE_DETAIL_PROFILE"].Row;

%% break footer
range := template.Range["RANGE_DETAIL_PROFILE"];
if !IsIDispatch(range) then
{
  errmsg := "Range RANGE_DETAIL_PROFILE not found!";
  goto error;
};
excel.Goto(range);
excel.Selection.EntireRow.Delete();
template.Calculate();

%% detail footer

nwffj_head:=4;
cpy_start:=nwffj_head;
cpy_end:=nwffj_head;
cpy_start1:=nwffj_head;
cpy_end1:=nwffj_head;
cpy_at:=14;


;

