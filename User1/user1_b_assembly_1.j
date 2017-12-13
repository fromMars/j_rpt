/* USER1_B_ASSEMBLY_1.J
 * show accessory prices */

/*序号*/
list_no_formula:="=row()-"+inttostr(row_increase+3);
recent_rowid:=-1;

; ******************************Estim Excel************************************
; *****************************************************************************
; %NAME% (%BATCH%)
; 


%% detail
; ******************************Estim Excel************************************
; *****************************************************************************
; %NAME% (%BATCH%) - Detail   b_assembly_1.j
; 

; Item price
RowId  := StrToNum(cList.Strings[bList.IndexOf("@%DB_COST_ARTICLE%"+"@%DB_COST_LOSSTYPE%")]);
if recent_rowid=-1 || recent_rowid>(rowid+row_increase) then
	recent_rowid:=rowid+row_increase;

CellCT := 'Indirect("Cost!"&address('+numtostr(strtonum(sList.Strings[cList.IndexOf(IntToStr(RowId))]))+","+IntToStr(ColCT)+"))/"+pList.Strings[cList.IndexOf(IntToStr(RowId))];
CellC1 := 'Indirect("Cost!"&address('+numtostr(strtonum(sList.Strings[cList.IndexOf(IntToStr(RowId))]))+","+IntToStr(ColC1)+"))";
CellC2 := 'Indirect("Cost!"&address('+numtostr(strtonum(sList.Strings[cList.IndexOf(IntToStr(RowId))]))+","+IntToStr(ColC2)+"))";
CellC7 := 'Indirect("Cost!"&address('+numtostr(strtonum(sList.Strings[cList.IndexOf(IntToStr(RowId))]))+","+IntToStr(ColC7)+"))";
CellC3 := 'Indirect("Cost!"&address('+numtostr(strtonum(sList.Strings[cList.IndexOf(IntToStr(RowId))]))+","+IntToStr(ColC3)+"))";
CellC4 := 'Indirect("Cost!"&address('+numtostr(strtonum(sList.Strings[cList.IndexOf(IntToStr(RowId))]))+","+IntToStr(ColC4)+"))";
CellC5 := 'Indirect("Cost!"&address('+numtostr(strtonum(sList.Strings[cList.IndexOf(IntToStr(RowId))]))+","+IntToStr(ColC5)+"))";
CellC6 := 'Indirect("Cost!"&address('+numtostr(strtonum(sList.Strings[cList.IndexOf(IntToStr(RowId))]))+","+IntToStr(ColC6)+"))";
if (StrToNum(StrReplace(pList.Strings[cList.IndexOf(IntToStr(RowId))],"%DECIMALSEP%","."),0) > 0) then
{
  TempValue   := StrReplace("@%DB_RES_PRICE%/%ASSEMBLYCOUNT%",".","%DECIMALSEP%");
  TempFormula := "=((((((((("+TempValue+")*("+CellCT+"))*(1+"+CellC1+"))*(1-"+CellC2+"))*"+CellC7+")*"+CellC3+")*(1+"+CellC6+"))*(1+"+CellC4+"))*(1-"+CellC5+"))";
  CurrentCell := CostSheet.Cells[RowId+row_increase][ColId];
  CurrentCell.Formula := TempFormula;
  CurrentCell.Font.Italic := False;
  CurrentCell.Interior.Color := Color;
  CurrentCell.Borders.LineStyle := 1;
}
else
{
  TempValue   := StrReplace("@%DB_RES_PRICE%/%ASSEMBLYCOUNT%",".","%DECIMALSEP%");
  TempFormula := "=(((((((("+TempValue+")*(1+"+CellC1+"))*(1-"+CellC2+"))*"+CellC7+")*"+CellC3+")*(1+"+CellC6+"))*(1+"+CellC4+"))*(1-"+CellC5+"))";
  CurrentCell := CostSheet.Cells[RowId+row_increase][ColId];
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

;序号
costsheet.cells[rowid+row_increase][1].formula:=list_no_formula;

;供应商
s_colid:=colid+1;
currentcell:=costsheet.cells[rowid+row_increase][s_colid];
currentcell.value:="%DSP_COST_SUPPLIER%";
if trim(currentcell.value)="EOSS" then
    currentcell.value:="易欧思专用";
currentcell.borders.linestyle:=1;

;单价
u_colid:=colid-1;
currentcell:=costsheet.cells[rowid+row_increase][u_colid];
u_recent_value:=currentcell.formula;
if "@%DB_COST_ASSEMBLY%"="" then
{
    /*工程级价格块*/
}
else
{
    /*计算单价，单价=金额/单樘用量*/
	tot_formula:="="+RId+LBr+inttostr(0)+RBr+CId+Lbr+"1"+RBr+"/"+RId+LBr+inttostr(0)+RBr+CId+Lbr+"-1"+RBr;
	currentcell.formulaR1C1:=tot_formula;
}
/*clear unused items*/
if @%DB_COST_ARTICLE%<>32 && @%DB_COST_ARTICLE%<>89 then
    currentcell.value:="";
currentcell.borders.linestyle:=1;

;单樘用量
wps_colid:=colid-2;
currentcell:=costsheet.cells[rowid+row_increase][wps_colid];
if "@%DB_COST_ASSEMBLY%"<>"" then
{
    currentcell.formulaR1C1:="=@COST_QUANTITY/%ASSEMBLYCOUNT%";
    /*clear unused items*/
    if @%DB_COST_ARTICLE%<>32 && @%DB_COST_ARTICLE%<>89 then
        currentcell.value:="";
    
}
else
{
    /*工程级价格块*/
    pla_formula:="="+numtostr(currentcell.value)+"*mianji/Cost!mianji";
    /*pla_formula:="=mianji";*/
	currentcell.formula:=pla_formula;
    currentcell_tmp:=costsheet.cells[rowid+row_increase][wps_colid+1];
    datasheet.range["HNDRate"].formula:='=Indirect("Cost!"&address('+inttostr(rowid)+","+inttostr(colid+7)+"))"+"/Cost!mianji";                 /*每平米用量*/
    currentcell_tmp.formula:="=Data!HNDRate";
    /*单樘用量*/
    currentcell_tmp:=costsheet.cells[rowid+row_increase][wps_colid+2];
    currentcell_tmp.formulaR1C1:="="+RId+CId+LBr+"-2"+RBr+"*"+RId+CId+LBr+"-1"+RBr;
    currentcell_tmp:=costsheet.cells[rowid+row_increase][wps_colid-2];
    currentcell_tmp.value:="㎡";
    currentcell_tmp.HorizontalAlignment:=-4108;
}
currentcell.borders.linestyle:=1;


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

tmp_rowid_increase:=RowId+row_increase;
RowId_1:=tmp_rowid_increase;
list_no_formula:="=row()-"+inttostr(row_increase+3);
