; ******************************Estim Excel************************************
; *****************************************************************************
; %NAME% (%BATCH%)
; 

ErrMsg := "";
oleserver := "excel.application";
goto begin;

; General exception raiser
@generalerror:
  if !excel.visible then excel.visible := true;
  MsgErr(FormatStr(GetLanText(-9688), ErrMsg));
  Halt;

; Error in oleserver
@oleservererror:
  if !excel.visible then excel.visible := true;
  MsgErr(FormatStr(GetLanText(-9687), ErrMsg));
  Halt;

; Start processing
@begin:
excel := start(oleserver);
if !IsIDispatch(excel) then
{
  ErrMsg := oleserver;
  goto oleservererror;
}
else
{
  oleversion := StrToNum(GetParam("OFFICE"), 0);
  oleversion := StrToNum(excel.Version, oleversion);
  excel.visible := True;
};

; Open a temporary template file for calculations
TemplateFile := FileSearch("%XLT_TEMPLATE%.XLT", "%PATH_DATA%");
if templatefile = "" || !FileExists(templatefile) then
{
  ErrMsg := "Cannot find template <%XLT_TEMPLATE%.XLT> in <%PATH_DATA%>!";
  goto generalerror;
}
Template := excel.workbooks.add(TemplateFile);
if !IsIDispatch(Template) then
{
  ErrMsg := "Open <" + TemplateFile + "> failed!";
  goto generalerror;
}
Template.Author := "%DB_USERDESC%";

if excel.worksheets.count>0 then
{
	curr_sheet:=template.worksheets[" ¶©»õÃ÷Ï¸µ¥"];
}
else
{
	ErrMsg:="no sheet found in "+TemplateFile+" file!";
	goto generalerror;
}

first_row:=curr_sheet.range["ProfList"];
rowid:=first_row.row;
init_rowid:=rowid;
;-------------------------------------------------------------------------------
%%detail
;-------------------------------------------------------------------------------

first_row.select();
excel.selection.entirerow.insert();
rowid:=first_row.row-2;
first_row.entirerow.copy();
curr_sheet.rows[rowid].entirerow.select();
curr_sheet.paste;

curr_cell:=curr_sheet.cells[rowid][1];
curr_cell.value:=(rowid-init_rowid+2)/2;

curr_cell:=curr_sheet.cells[rowid][2];
curr_cell.value:="%DSP_PIECE_PROFILEDESC%";

curr_cell:=curr_sheet.cells[rowid][3];
curr_cell.value:="%DSP_PIECE_PRODUCT%";

curr_cell:=curr_sheet.cells[rowid][5];
curr_cell.value:="=@%DB_PIECE_LOPT%/1000";

curr_cell:=curr_sheet.cells[rowid][6];
curr_cell.value:=%DSP_PIECE_FACTOR%;

/*
curr_cell:=curr_sheet.cells[rowid]8];
curr_cell.value:=rowid-init_rowid-1;

curr_cell:=curr_sheet.cells[rowid][9];
curr_cell.value:=rowid-init_rowid-1;*/


;-------------------------------------------------------------------------------
%% detail footer
;-------------------------------------------------------------------------------

curr_sheet.usedrange.rows[""+inttostr(init_rowid)+":"+inttostr(rowid)].borders.linestyle:=1;
