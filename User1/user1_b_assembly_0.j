/* USER1_B_ASSEMBLY_0.J
 * show profile prices */


Count:=recent_count;
ColId:=recent_colid;                                            /*14 单价列*/
CostSheet:=recent_cost_sheet;                                   /*sheet "COST"*/

CostSheet.Copy(CostSheet);                                      /*复制sheet*/
cnt := Template.WorkSheets.Count-3;

template_cost:=Template.WorkSheets[cnt];
template_cost.Name := "Cost_"+trim("%ASSEMBLY_TEXT%");
template_cost.Activate();

CostSheet:=template_cost;

CostSheet.Range["RateRows"].Delete;
ColId:=ColId-8;


recent_rowid:=-1;                                               /*记录小计行前最后一行位置*/


curr_assembly:=getcurrentproject().projectdata.currentassembly; /*获取Editor中窗组对象*/
assembly_cnt:=getcurrentproject().projectdata.childcount;       /*子对象数量*/
i_cnt:=0;
/*获取当前计算窗组对象*/
while i_cnt<assembly_cnt do
{
    curr_assembly:=getcurrentproject().projectdata.children[i_cnt];
    if curr_assembly.code="%ASSEMBLY_TEXT%" then
        break;
    i_cnt:=i_cnt+1;
}

/*窗组名称*/
/*curr_name:=curr_assembly.code;*/
costsheet.range["chuanghao"].value:=curr_assembly.code;

/*面积*/
f_cnt:=0;
curr_frame:=curr_assembly.children[0];
frame_cnt:=curr_assembly.childcount;
a_mianji:=0;
/*计算同一窗组各Frame的尺寸，其和为总面积*/
while f_cnt<frame_cnt do
{
    curr_frame:=curr_assembly.children[f_cnt];
    f_width:=curr_frame.width;
    f_height:=curr_frame.height;
    if f_cnt=0 then
    {
        helpsheet.cells[help_cnt][4].value:=f_width;            /*helpsheet主框尺寸_宽*/
        helpsheet.cells[help_cnt][5].value:=f_height;           /*helpsheet主框尺寸_高*/
    }
    if f_cnt>0 then
    {
        helpsheet.cells[help_cnt][4].value:='';                  /*门联窗*/
        helpsheet.cells[help_cnt][5].value:='';                  /*不填写*/
    }
    f_mianji:=f_width*f_height;
    a_mianji:=a_mianji+f_mianji;
    f_cnt:=f_cnt+1;
}
curr_surface:=a_mianji/1000000;
costsheet.range["mianji"].value:=curr_surface;
costsheet.range["mianji"].HorizontalAlignment:=-4131;
costsheet.range["mianji"].offset[0][-1].HorizontalAlignment:=-4152;

a_fee_row:=0;

/*used to calculate A*/
RowId_0:=0;
RowId_1:=0;
RowId_2:=0;
RowId_A:=0;

/*工程总面积，用以计算工程级价格块*/
total_area:=total_area+curr_surface*%ASSEMBLYCOUNT%;
cost_ori:=template.worksheets["cost"];
cost_ori.range["mianji"].value:=total_area;

/*窗型及樘数，用于XXX_E_PROJECT.J计算*/
/*上文计算面积过程中已输出其主框尺寸的宽高值*/
helpsheet.cells[help_cnt][3].value:='';                           /*窗型*/
helpsheet.cells[help_cnt][6].value:="%ASSEMBLYCOUNT%";             /*樘数*/
help_cnt:=help_cnt+1;

; ******************************Estim Excel************************************
; *****************************************************************************
; %NAME% (%BATCH%)  b_assembly_0.j
; 

colid:=6;                                                   /*单价列*/
Count := Count + 1;
ColId := ColId + 1;                                         /*金额列*/
Color := DataSheet.Range["CellFormat"].Interior.Color;

/*初始化窗组价格*/
i := 0;
while (i < cList.Count-3) do
{
  RowId       := StrToNum(cList.Strings[i]);
  TempValue   := 0.0;
  CurrentCell := CostSheet.Cells[RowId][ColId];             /*金额列*/
  CurrentCell.Value := TempValue;
  currentcell.NumberFormat:=CellCostFormat;
  CurrentCell1 := CostSheet.Cells[RowId][ColId-1];          /*单价列*/
  if CurrentCell1.Value=0 then
    CurrentCell1.Value := TempValue;
  currentcell1.NumberFormat:=CellCostFormat;
  CurrentCell0 := CostSheet.Cells[RowId][ColId-2];          /*单樘用量列*/
  if CurrentCell0.value=0 then
    CurrentCell0.Value := TempValue;
  CurrentCell.Font.Italic := True;
  CurrentCell.Interior.Color := Color;
  CurrentCell.Borders.LineStyle := 1;
  i := i + 1;
};

/*calculate follow artikels, recent_profile_value-recent TempValue[string],tmp_tmp_value-current TempValue[string]*/
/*计算跟随价格块*/
recent_profile_value:="0";
tmp_tmp_value:="0";

a_linked:=strings.create();                                 /*记录跟随价格块*/

%% detail
; ******************************Estim Excel************************************
; *****************************************************************************
; %NAME% (%BATCH%) - Detail
; 

a_link:="";
z_pg:=pricegroups.create();
z_pg.code.group:="A";
z_pg.code.block:=@%DB_COST_ARTICLE%;
if z_pg.find() then
{
	a_link:=z_pg.link;
    if a_link<>"" && a_linked.indexof(a_link)=-1 then       /*跟随价格块存在且未记录*/
        a_linked.add(a_link);                               /*则记录该价格块*/
}
else
{
	msgbox("no article block "+inttostr(z_pg.code.block)+" found!");
}

/*获取行号*/
if a_link<>"" then
{
	RowId  := StrToNum(cList.Strings[bList.IndexOf(a_link+"@%DB_PIECE_LOSSTYPE%")]);
}
else
{
	RowId  := StrToNum(cList.Strings[bList.IndexOf("@%DB_PIECE_ARTICLE%"+"@%DB_PIECE_LOSSTYPE%")]);
}


if recent_rowid=-1 || recent_rowid>rowid then
	recent_rowid:=rowid;

CellCT := 'Indirect("Cost!"&address('+sList.Strings[cList.IndexOf(IntToStr(RowId))]+","+IntToStr(ColCT)+"))/"+pList.Strings[cList.IndexOf(IntToStr(RowId))];
CellC1 := 'Indirect("Cost!"&address('+sList.Strings[cList.IndexOf(IntToStr(RowId))]+","+IntToStr(ColC1)+"))";
CellC2 := 'Indirect("Cost!"&address('+sList.Strings[cList.IndexOf(IntToStr(RowId))]+","+IntToStr(ColC2)+"))";
CellC7 := 'Indirect("Cost!"&address('+sList.Strings[cList.IndexOf(IntToStr(RowId))]+","+IntToStr(ColC7)+"))";
CellC3 := 'Indirect("Cost!"&address('+sList.Strings[cList.IndexOf(IntToStr(RowId))]+","+IntToStr(ColC3)+"))";
CellC4 := 'Indirect("Cost!"&address('+sList.Strings[cList.IndexOf(IntToStr(RowId))]+","+IntToStr(ColC4)+"))";
CellC5 := 'Indirect("Cost!"&address('+sList.Strings[cList.IndexOf(IntToStr(RowId))]+","+IntToStr(ColC5)+"))";
CellC6 := 'Indirect("Cost!"&address('+sList.Strings[cList.IndexOf(IntToStr(RowId))]+","+IntToStr(ColC6)+"))";
if (StrToNum(StrReplace(pList.Strings[cList.IndexOf(IntToStr(RowId))],"%DECIMALSEP%","."),0) > 0) then
{
  CurrentCell := CostSheet.Cells[RowId][ColId];
  TempValue   := StrReplace("@%DB_PIECE_PRICE%/%ASSEMBLYCOUNT%",".","%DECIMALSEP%");
  /*当前价格块跟随或当前价格块被跟随-true，否则-false*/
  if a_link<>"" || a_linked.indexof("@%DB_COST_ARTICLE%")<>-1 then
  {
    /*为current_profile_value保存当前tempvalue(recent_profile_value)*/
	tmp_tmp_value:=tempvalue;
	curr_profile_value:="((((((((("+recent_profile_value+")*("+CellCT+"))*(1+"+CellC1+"))*(1-"+CellC2+"))*"+CellC7+")*"+CellC3+")*(1+"+CellC6+"))*(1+"+CellC4+"))*(1-"+CellC5+"))";
	TempFormula := "=((((((((("+TempValue+")*("+CellCT+"))*(1+"+CellC1+"))*(1-"+CellC2+"))*"+CellC7+")*"+CellC3+")*(1+"+CellC6+"))*(1+"+CellC4+"))*(1-"+CellC5+"))+"+curr_profile_value;       /*增加curr_profile_value*/
  }
  else
  {
	TempFormula := "=((((((((("+TempValue+")*("+CellCT+"))*(1+"+CellC1+"))*(1-"+CellC2+"))*"+CellC7+")*"+CellC3+")*(1+"+CellC6+"))*(1+"+CellC4+"))*(1-"+CellC5+"))";
  }
  CurrentCell.Formula := TempFormula;
  CurrentCell.Interior.Color := Color;
  CurrentCell.Borders.LineStyle := 1;
}
else
{
  CurrentCell := CostSheet.Cells[RowId][ColId];
  TempValue   := StrReplace("@%DB_PIECE_PRICE%/%ASSEMBLYCOUNT%",".","%DECIMALSEP%");
  if a_link<>"" || a_linked.indexof("@%DB_COST_ARTICLE%")<>-1 then
  {
	tmp_tmp_value:=tempvalue;
	curr_profile_value:="(((((((("+recent_profile_value+")*(1+"+CellC1+"))*(1-"+CellC2+"))*"+CellC7+")*"+CellC3+")*(1+"+CellC6+"))*(1+"+CellC4+"))*(1-"+CellC5+"))";
	TempFormula := "=(((((((("+TempValue+")*(1+"+CellC1+"))*(1-"+CellC2+"))*"+CellC7+")*"+CellC3+")*(1+"+CellC6+"))*(1+"+CellC4+"))*(1-"+CellC5+"))+"+curr_profile_value;
  }
  else
  {
	TempFormula := "=(((((((("+TempValue+")*(1+"+CellC1+"))*(1-"+CellC2+"))*"+CellC7+")*"+CellC3+")*(1+"+CellC6+"))*(1+"+CellC4+"))*(1-"+CellC5+"))";
  }
  CurrentCell.Formula := TempFormula;
  CurrentCell.Font.Italic := False;
  CurrentCell.Interior.Color := Color;
  CurrentCell.Borders.LineStyle := 1;
}


; Item formula
/*
TempFormula := '=Indirect(address('+sList.Strings[cList.IndexOf(IntToStr(RowId))]+','+IntToStr(ColId)+',,,"Cost"))*Indirect(address('+sList.Strings[bList.IndexOf("-2")]+','+IntToStr(ColId)+',,,"Cost"))';
CurrentCell := HelpSheet.Cells[RowId][ColId];
CurrentCell.Formula := TempFormula;
CurrentCell.Font.Italic := %IF{@%DB_RES_PRICE%,False,True};
CurrentCell.Interior.Color := Color;
CurrentCell.Borders.LineStyle := 1;*/



;型材单位
un_colid:=3;
currentcell:=costsheet.cells[rowid][un_colid];
currentcell.value:="kg";
currentcell.HorizontalAlignment:=-4108;


;供应商
s_colid:=colid+1;
currentcell:=costsheet.cells[rowid][s_colid];
currentcell.value:="%DSP_PIECE_SUPPLIER%";
if trim(currentcell.value)="EOSS" then
    currentcell.value:="华建易欧思"+trim("%DSP_PIECE_SYSTEM%")+"系列型材";
currentcell.borders.linestyle:=1;


;单位面积重量
wps_colid:=colid-2;
currentcell:=costsheet.cells[rowid][wps_colid];
if a_link<>"" || a_linked.indexof("@%DB_COST_ARTICLE%")<>-1 then
{
	curr_profile_value:=currentcell.value;
    currentcell.formulaR1C1:="=@%DB_PIECE_WEIGHT%/%ASSEMBLYCOUNT%+"+numtostr(curr_profile_value);
	currentcell.borders.linestyle:=1;
}
else
{
    currentcell.formulaR1C1:="=@%DB_PIECE_WEIGHT%/%ASSEMBLYCOUNT%";
	currentcell.borders.linestyle:=1;
}


;单价
u_colid:=colid-1;
currentcell:=costsheet.cells[rowid][u_colid];
u_recent_value:=currentcell.value;
if u_recent_value<>0 && a_link="" && a_linked.indexof("@%DB_COST_ARTICLE%")=-1 then
{
	costsheet.cells[rowid][u_colid+1].value:=u_recent_value;
	currentcell.value:="";
}
else
{
	if a_link<>"" || a_linked.indexof("@%DB_COST_ARTICLE%")<>-1 then
	{
		curr_profile_value:=currentcell.value;
        tot_formula:="="+RId+CId+LBr+"1"+RBr+"/"+RId+CId+LBr+"-1"+RBr;
		currentcell.formulaR1C1:=tot_formula;
	}
	else
	{
		tot_formula:="="+RId+CId+Lbr+"1"+RBr+"/(@%DB_PIECE_WEIGHT%/%ASSEMBLYCOUNT%)";
		currentcell.formulaR1C1:=tot_formula;
	}
}
currentcell.borders.linestyle:=1;

/*跟随价格块计算用，本次tempvalue保存在recent_profile_value中，下次计算用*/
recent_profile_value:=tmp_tmp_value;
tmp_tmp_value:="0";

%% break header
; ******************************Estim Excel************************************
; *****************************************************************************
; %NAME% (%BATCH%) - Break header
; 

%% break footer
; ******************************Estim Excel************************************
; *****************************************************************************
; %NAME% (%BATCH%) - Break footer
; 

%% detail footer
; ******************************Estim Excel************************************
; *****************************************************************************
; %NAME% (%BATCH%) - Detail footer
; 


/*新增小计行，rowid仍为最后一项所在行，row_increase为新增行数*/
CostSheet.Rows[RowId+1].select();
excel.Selection.EntireRow.Insert();
excel.Selection.EntireRow.Insert();
row_increase:=2;

CostSheet.Range[CostSheet.Cells[RowId+1][1]][CostSheet.Cells[RowId+2][1]].merge();
CostSheet.Cells[RowId+1][1].Value:="小计";
costsheet.cells[rowid+1][1].VerticalAlignment:=-4108;
costsheet.cells[rowid+1][1].HorizontalAlignment:=-4108;
CostSheet.Range[CostSheet.Cells[RowId+1][2]][CostSheet.Cells[RowId+1][3]].merge();
CostSheet.Cells[RowId+1][2].Value:="型材损耗";
CostSheet.Range[CostSheet.Cells[RowId+2][2]][CostSheet.Cells[RowId+2][3]].merge();
CostSheet.Cells[RowId+2][2].Value:="型材小计";
CostSheet.Range[CostSheet.Cells[RowId+1][5]][CostSheet.Cells[RowId+1][7]].merge();

/*型材损耗*/
CostSheet.Cells[RowId+1][5].value:=0;
CostSheet.Cells[RowId+1][5].NumberFormatLocal:="0.0%";
CostSheet.Range[CostSheet.Cells[RowId+2][5]][CostSheet.Cells[RowId+2][7]].merge();
costsheet.cells[rowid+2][5].NumberFormat:=CellCostFormat;

/*型材小计*/
Formula1 := "="+SumFormulaText+"("+RId+LBr+IntToStr(recent_rowid-rowid-2)+RBr+CId+LBr+"2"+RBr+":"+RId+LBr+"-2"+RBr+CId+Lbr+"2"+RBr+")*(1+"+RId+LBr+"-1"+RBr+CId+")";
CostSheet.Cells[RowId+2][5].FormulaR1C1:=Formula1;

/*背景色*/
CostSheet.Range[CostSheet.Cells[RowId+1][1]][CostSheet.Cells[RowId+1][8]].Interior.Color:=14935011;
CostSheet.Range[CostSheet.Cells[RowId+2][1]][CostSheet.Cells[RowId+2][8]].Interior.Color:=14935011;

Rowid_0:=RowId+2;


