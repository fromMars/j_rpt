;order_list_per_supplier_b_project_2
;Bestellijst/leverancier: Beslag

%% detail
range := template.Range["RANGE_DETAIL_ACC"];
if !IsIDispatch(range) then
{
  errmsg := "Range RANGE_DETAIL_ACC not found!";
  goto error;
};
excel.Goto(range);
row := row + 1;
range := range.Offset[row];
excel.Goto(range);
excel.Selection.EntireRow.Insert();
template.Range["RANGE_DETAIL_ACC"].Copy();
template.Paste;
range := range.Offset[-1];
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
  cell := range.Cells[1][6];
  if IsIDispatch(cell) then cell.Value := @%DB_ATTRIB_ITMPRICE%;
  cell := range.Cells[1][7];
  if IsIDispatch(cell) then cell.Value := @%DB_ATTRIB_MINSIZE%;
  cell := range.Cells[1][8];
  if IsIDispatch(cell) then cell.Value := @%DB_ATTRIB_QTYPRICE%;
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
  if IsIDispatch(cell) then cell.Value := trim("not defined");
  cell := range.Cells[1][7];
  if IsIDispatch(cell) then cell.Value := @%DB_ATTRIB_PACKCOUNT%;
  cell := range.Cells[1][8];
  if IsIDispatch(cell) then cell.Value := @%DB_ATTRIB_PRICE%;
}

/*cpy_range := book.WorkSheets[1].Range("A" & cpy_start);*/

/*excel.Goto(range);
excel.Selection.EntireRow.Copy();*/

/*
book.WorkSheets[1].Activate();
book.WorkSheets[1].Range["gcnzfjb_line"].Copy();
excel.Goto(book.WorkSheets[1].Range["A" + Inttostr(cpy_end+1)]);
excel.Selection.EntireRow.Insert();
book.WorkSheets[1].Paste;

cpy_end:=cpy_start+row;
book.WorkSheets[1].Cells[cpy_end][1].Value:=cpy_end-4;
book.WorkSheets[1].Cells[cpy_end][2].Value:=iif(trim("%DSP_ATTRIB_ARTICLECODE%")="",trim("%DSP_ATTRIB_ACC%")+"."+trim(%IF{%ORDER_LIST_PER_SUPPLIER_VARIANT%=0,"%DSP_ATTRIB_SERIE%","%DSP_ATTRIB_VARIETY%"}),trim("%DSP_ATTRIB_ARTICLECODE%"));
cell := range.Cells[1][3];
book.WorkSheets[1].Cells[cpy_end][4].Value:=HtmlToNormalStr(trim("%DSP_ATTRIB_VARIETYDESC%"));
book.WorkSheets[1].Cells[cpy_end][3].Value:=HtmlToNormalStr(trim("%DSP_ATTRIB_ACCDESC%"));
book.WorkSheets[1].Cells[cpy_end][5].Value:=@%DB_ATTRIB_NO%;
/////book.WorkSheets[1].Cells[cpy_end][6].Value:=@%DB_ATTRIB_PACKCOUNT%;
book.WorkSheets[1].Cells[cpy_end][7].Value:=@%DB_ATTRIB_PRICE%;
book.WorkSheets[1].Cells[cpy_end][8].Value:=trim("%DSP_ATTRIB_SUPPLIER%");*/

cpy_end:=cpy_start+row;


%% break header
tmp_acc.Copy(tmp_profile);
cnt := book.WorkSheets.Count-3;
if cnt <= 0 then
{
  errmsg := "No of sheets in base template is incorrect!";
  goto error;
};
template := book.WorkSheets[cnt];
if !IsIDispatch(template) then
{
  errmsg := "Destination sheet not found!";
  goto error;
};
template.Name := "acc_"+trim("%DSP_ATTRIB_SUPPLIER%");
template.Activate();
cell := template.Range["CELL_ACC_SUPPLIER"];
if IsIDispatch(cell) then cell.Value := "%DSP_TEXT_CLIENT%";
cell := template.Range["CELL_ACC_SUPPLIERDESC"];
if IsIDispatch(cell) then cell.Value := "%DSP_TEXT_CONTACT%";
cell := template.Range["CELL_ACC_SUPPLIERADDRESS1"];
if IsIDispatch(cell) then cell.Value := "%DSP_TEXT_STREET%";
cell := template.Range["CELL_ACC_SUPPLIERADDRESS2"];
if IsIDispatch(cell) then cell.Value := trim("%DSP_TEXT_ZIP%")+" "+trim("%DSP_TEXT_PLACE%");
cell := template.Range["CELL_ACC_SUPPLIERADDRESS3"];
if IsIDispatch(cell) then cell.Value := "%DSP_TEXT_COUNTRY%";
cell := template.Range["CELL_ACC_SUPPLIERPHONE"];
if IsIDispatch(cell) then cell.Value := "%DSP_TEXT_PHONE%";
cell := template.Range["CELL_ACC_SUPPLIERFAX"];
if IsIDispatch(cell) then cell.Value := "%DSP_TEXT_TELEFAX%";

; Translation of header price labels.
cell := template.Range["PURCHASE_PRICE_LABEL2"];
if IsIDispatch(cell) then cell.Value := GetLanText(-312) + " (" + "%CURRENCY%" +")";
cell := template.Range["UNIT_PRICE_LABEL2"];
if IsIDispatch(cell) then cell.Value := GetLanText(-26005) + " (" + "%CURRENCY%" +")";
cell := template.Range["GRADUATE_PRICE_LABEL2"];
if IsIDispatch(cell) then cell.Value := GetLanText(-26006) + " (" + "%CURRENCY%" +")";
cell := template.Range["PACKING_PRICE_LABEL2"];
if IsIDispatch(cell) then cell.Value := GetLanText(-26007) + " (" + "%CURRENCY%" +")";

row := 0;
start := template.Range["RANGE_DETAIL_ACC"].Row;

/*
cpy_sheet:="增加";
cpy_start:=cpy_end;*/
cpy_end:=0;
cpy_start:=start;
p_start:=4;


tmp_nwffj.Copy(tmp_profile);
cnt := book.WorkSheets.Count-3;
if cnt <= 0 then
{
  errmsg := "No of sheets in base template is incorrect!";
  goto error;
};
template_nw := book.WorkSheets[cnt];
if !IsIDispatch(template_nw) then
{
  errmsg := "Destination sheet not found!";
  goto error;
};
template_nw.Name := "nwffj_"+trim("%DSP_ATTRIB_SUPPLIER%");
if trim("%DSP_ATTRIB_SUPPLIER%")="NF" then
	template_nw.Range["A1"].Value:=trim("工程内装附件表");
else if trim("%DSP_ATTRIB_SUPPLIER%")="WF" then
	template_nw.Range["A1"].Value:=trim("工程外发附件表");
else
	template_nw.Range["A1"].Value:=trim("%DSP_ATTRIB_SUPPLIER%");
template_nw.Range["C2"].Value:=trim("%PROJECT%");


%% break footer

range := template.Range["RANGE_DETAIL_ACC"];
if !IsIDispatch(range) then
{
  errmsg := "Range RANGE_DETAIL_ACC not found!";
  goto error;
}
excel.Goto(range);
excel.Selection.EntireRow.Delete();
template.Calculate();


template.Activate();
cpy_at:=start;
excel.Range["M"+inttostr(cpy_at)+":"+"AB"+inttostr(cpy_end)].Select();
/*excel.Rows[inttostr(cpy_start)+":"+inttostr(cpy_end)].Select();*/
excel.Selection.Copy();
excel.Selection.PasteSpecial(-4163);

template_nw.Activate();
/*book.WorkSheets[1].Range["gcnzfjb_line"].Copy();*/
excel.Goto(template_nw.Range["A" + Inttostr(p_start)]);
/*excel.Selection.EntireRow.Insert();*/
template_nw.Paste;

p_start:=cpy_end+1;

/*
p_end:=template_nw.UsedRange.Rows.Count;
no_cnt:=1;
no_cnt1:=4;
while no_cnt1<=p_end do
{
	template_nw.Range["A"+inttostr(no_cnt1)].Value:=no_cnt;
	no_cnt:=no_cnt+1;
	no_cnt1:=no_cnt1+1;
}
*/
%% detail footer
;

%%
;