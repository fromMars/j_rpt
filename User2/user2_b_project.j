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
	curr_sheet:=template.worksheets[" 订货明细单"];
}
else
{
	ErrMsg:="no sheet found in "+TemplateFile+" file!";
	goto generalerror;
}

curr_sheet.range["ProjectName"].value:=" 工程名称：%DSP_PIECE_PROJECT%";

first_row:=curr_sheet.range["ProfList"];
rowid:=first_row.row;
init_rowid:=rowid;

seperated_profile:=profiles.create();

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


seperated_profile.code.system:="@%DB_PIECE_SYSTEM%";
seperated_profile.code.profile:="@%DB_PIECE_PROFILE%";
inside_profile:="";
outside_profile:="";

if !seperated_profile.find() then
	msgbox("profile not found in profile.db!");
else
{
	seperated_cnt:=0;
	while seperated_cnt<5 do
	{
		if seperated_profile.accessories[seperated_cnt].colour=2 then
			inside_profile:=trim(seperated_profile.accessories[seperated_cnt].code.code);
		else if seperated_profile.accessories[seperated_cnt].colour=1 then
			outside_profile:=trim(seperated_profile.accessories[seperated_cnt].code.code);
		seperated_cnt:=seperated_cnt+1;
	}
}



if "@%DB_PIECE_INSIDE%"<>"" && "@%DB_PIECE_OUTSIDE%"<>"" then
{
    if strpos("_",inside_profile)=1 then
        inside_profile:=strdeletel(inside_profile,1);
    if strpos("_",outside_profile)=1 then
        outside_profile:=strdeletel(outside_profile,1);
    
	curr_cell:=curr_sheet.cells[rowid][8];
	curr_cell.value:=inside_profile;
	curr_cell:=curr_sheet.cells[rowid+1][8];
	curr_cell.value:="@%DB_PIECE_INSIDE%";
	
	curr_cell:=curr_sheet.cells[rowid][9];
	curr_cell.value:=outside_profile;
	curr_cell:=curr_sheet.cells[rowid+1][9];
	curr_cell.value:="@%DB_PIECE_OUTSIDE%";
}
else
{
	curr_sheet.range[curr_sheet.cells[rowid][8]][curr_sheet.cells[rowid+1][9]].merge();
	curr_cell:=curr_sheet.cells[rowid][8].value:="%DSP_PIECE_SERIE%";
}

;-------------------------------------------------------------------------------
%% detail footer
;-------------------------------------------------------------------------------

if %GLOBAL_PRICE_PROFILE%=1 then
	curr_sheet.cells[init_rowid][10].value:="易菲特隔热条"+chr(10)+"超高精级";
else
	curr_sheet.cells[init_rowid][10].value:="泰诺风隔热条"+chr(10)+"超高精级";
curr_sheet.range[curr_sheet.cells[init_rowid][10]][curr_sheet.cells[rowid][10]].merge();
curr_sheet.usedrange.rows[""+inttostr(init_rowid)+":"+inttostr(rowid+1)].borders.linestyle:=1;



