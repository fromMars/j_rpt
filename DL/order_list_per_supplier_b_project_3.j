;order_list_per_supplier_b_project_3
;Bestellijst/leverancier: Rubbers

/*
excel.Goto(book.WorkSheets[1].Range["A1:J3"]);
excel.Selection.EntireRow.Copy();
excel.Goto(book.WorkSheets[1].Range["A"&cpy_start]);
excel.Selection.EntireRow.Paste;*/

%% detail
range := template.Range["RANGE_DETAIL_GASKET"];
if !IsIDispatch(range) then
{
  errmsg := "Range RANGE_DETAIL_GASKET not found!";
  goto error;
};
excel.Goto(range);
row := row + 1;
range := range.Offset[row];
if !IsIDispatch(range) then
{
  errmsg := "Range RANGE_DETAIL_GASKET is incorrect!";
  goto error;
};
excel.Goto(range);
excel.Selection.EntireRow.Insert();
template.Range["RANGE_DETAIL_GASKET"].Copy();
template.Paste;
range := range.Offset[-1];
if !IsIDispatch(range) then
{
  errmsg := "Range RANGE_DETAIL_GASKET is incorrect!";
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
;;;;;;;;;;book.WorkSheets[1].Cells[cpy_end][6].Value:=@%DB_ATTRIB_PACKCOUNT%;
book.WorkSheets[1].Cells[cpy_end][7].Value:=@%DB_ATTRIB_PRICE%;
book.WorkSheets[1].Cells[cpy_end][8].Value:=trim("%DSP_ATTRIB_SUPPLIER%");
*/

cpy_end:=cpy_start+row;


%% break header
; dupliceer sheet "gasket" en maak hoofding met volgende info:
tmp_gasket.Copy(tmp_profile);
cnt := book.WorkSheets.Count-3;
if cnt <= 0 then
{
  errmsg := "No of sheets in base template is incorrect!";
  goto error;
};
template := book.WorkSheets[cnt];
template.Name := "gasket_"+trim("%DSP_ATTRIB_SUPPLIER%");
template.Activate();
cell := template.Range["CELL_GASKET_SUPPLIER"];
if IsIDispatch(cell) then cell.Value := "%DSP_TEXT_CLIENT%";
cell := template.Range["CELL_GASKET_SUPPLIERDESC"];
if IsIDispatch(cell) then cell.Value := "%DSP_TEXT_CONTACT%";
cell := template.Range["CELL_GASKET_SUPPLIERADDRESS1"];
if IsIDispatch(cell) then cell.Value := "%DSP_TEXT_STREET%";
cell := template.Range["CELL_GASKET_SUPPLIERADDRESS2"];
if IsIDispatch(cell) then cell.Value := trim("%DSP_TEXT_ZIP%")+" "+trim("%DSP_TEXT_PLACE%");
cell := template.Range["CELL_GASKET_SUPPLIERADDRESS3"];
if IsIDispatch(cell) then cell.Value := "%DSP_TEXT_COUNTRY%";
cell := template.Range["CELL_GASKET_SUPPLIERPHONE"];
if IsIDispatch(cell) then cell.Value := "%DSP_TEXT_PHONE%";
cell := template.Range["CELL_GASKET_SUPPLIERFAX"];
if IsIDispatch(cell) then cell.Value := "%DSP_TEXT_TELEFAX%";

; Translation of header price labels.
cell := template.Range["PURCHASE_PRICE_LABEL3"];
if IsIDispatch(cell) then cell.Value := GetLanText(-312) + " (" + "%CURRENCY%" +")";
cell := template.Range["UNIT_PRICE_LABEL3"];
if IsIDispatch(cell) then cell.Value := GetLanText(-26005) + " (" + "%CURRENCY%" +")";
cell := template.Range["GRADUATE_PRICE_LABEL3"];
if IsIDispatch(cell) then cell.Value := GetLanText(-26006) + " (" + "%CURRENCY%" +")";
cell := template.Range["PACKING_PRICE_LABEL3"];
if IsIDispatch(cell) then cell.Value := GetLanText(-26007) + " (" + "%CURRENCY%" +")";

row := 0;
start := template.Range["RANGE_DETAIL_GASKET"].Row;

/*
cpy_sheet:="Ôö¼Ó";
cpy_start:=cpy_end;*/
cpy_end:=0;
cpy_start:=start;

ws_exists:=0;
ws_exists_cnt:=cnt;
while ws_exists_cnt>0 do
{
	if book.Worksheets[ws_exists_cnt].Name="nwffj_"+trim("%DSP_ATTRIB_SUPPLIER%") then
	{
		ws_exists:=1;
		break;
	}
	ws_exists_cnt:=ws_exists_cnt-1;
}

if ws_exists=1 then
{
	template_nw:=book.WorkSheets["nwffj_"+trim("%DSP_ATTRIB_SUPPLIER%")];
}
else
{
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
}
p_start:=template_nw.UsedRange.Rows.Count;

%% break footer
/*break footer begins*/
range := template.Range["RANGE_DETAIL_GASKET"];
if !IsIDispatch(range) then
{
  errmsg := "Range RANGE_DETAIL_GASKET not found!";
  goto error;
};
excel.Goto(range);
excel.Selection.EntireRow.Delete();
template.Calculate();



/*df here recently*/
/*
excel.Goto(book.WorkSheets[1].Range["gcnzfjb_line"]);
excel.Selection.EntireRow.Delete();*/



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



%% detail footer
;

%%
;

ws_cnt:=book.WorkSheets.Count;
ws_cnt1:=1;
while ws_cnt1<=ws_cnt do
{
	tmp_str:=book.WorkSheets[ws_cnt1].Name;
	if SubStr(tmp_str,1,6)="nwffj_" then
	{
		book.WorkSheets[ws_cnt1].Activate();
		p_end:=book.WorkSheets[ws_cnt1].UsedRange.Rows.Count;
		no_cnt:=1;
		no_cnt1:=nwffj_head;
		
		while no_cnt1<=p_end do
		{
			book.WorkSheets[ws_cnt1].Range["A"+inttostr(no_cnt1)].Value:=no_cnt;
			no_cnt:=no_cnt+1;
			no_cnt1:=no_cnt1+1;
		}
		book.WorkSheets[ws_cnt1].Range["C"+inttostr(nwffj_head)+":C"+inttostr(p_end)].Select();
		excel.Selection.Copy();
		excel.Goto(book.WorkSheets[ws_cnt1].Range["Z"+inttostr(nwffj_head)]);
		book.WorkSheets[ws_cnt1].Paste;
		book.WorkSheets[ws_cnt1].Range["D"+inttostr(nwffj_head)+":D"+inttostr(p_end)].Select();
		excel.Selection.Copy();
		excel.Goto(book.WorkSheets[ws_cnt1].Range["C"+inttostr(nwffj_head)]);
		book.WorkSheets[ws_cnt1].Paste;
		book.WorkSheets[ws_cnt1].Range["G"+inttostr(nwffj_head)+":G"+inttostr(p_end)].Select();
		excel.Selection.Copy();
		excel.Goto(book.WorkSheets[ws_cnt1].Range["E"+inttostr(nwffj_head)]);
		book.WorkSheets[ws_cnt1].Paste;
		book.WorkSheets[ws_cnt1].Range["H"+inttostr(nwffj_head)+":H"+inttostr(p_end)].Select();
		excel.Selection.Copy();
		excel.Goto(book.WorkSheets[ws_cnt1].Range["G"+inttostr(nwffj_head)]);
		book.WorkSheets[ws_cnt1].Paste;
		book.WorkSheets[ws_cnt1].Range["Z"+inttostr(nwffj_head)+":Z"+inttostr(p_end)].Select();
		excel.Selection.Copy();
		excel.Goto(book.WorkSheets[ws_cnt1].Range["D"+inttostr(nwffj_head)]);
		book.WorkSheets[ws_cnt1].Paste;
		
		
		book.WorkSheets[ws_cnt1].Range["H"+inttostr(nwffj_head)+":Z"+inttostr(p_end)].Select();
		excel.Selection.Delete();

		book.WorkSheets[ws_cnt1].Rows[3].Copy();
		book.WorkSheets[ws_cnt1].Rows[inttostr(nwffj_head)+":"+inttostr(p_end)].Select();
		excel.Selection.PasteSpecial(-4122);
		excel.Selection.Columns.AutoFit();
		
	}
	ws_cnt1:=ws_cnt1+1;
	
}

/*break footer ends*/

