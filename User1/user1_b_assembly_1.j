
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
/*RowId  := StrToNum(cList.Strings[bList.IndexOf("@%DB_COST_ARTICLE%"+"@%DB_COST_LOSSTYPE%"+"@%DB_COST_RATION%"+"@%DB_COST_FACTOR%"+"@%DB_COST_RATIO%")]);*/
RowId  := StrToNum(cList.Strings[bList.IndexOf("@%DB_COST_ARTICLE%"+"@%DB_COST_LOSSTYPE%")]);
if recent_rowid=-1 then
	recent_rowid:=rowid+row_increase;

CellCT := 'Indirect(Cost!address('+numtostr(strtonum(sList.Strings[cList.IndexOf(IntToStr(RowId))]))+","+IntToStr(ColCT)+"))/"+pList.Strings[cList.IndexOf(IntToStr(RowId))];
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
/*  CurrentCell.NumberFormat := CellPriceFormat;*/
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
/*  CurrentCell.NumberFormat := CellPriceFormat;*/
  CurrentCell.Font.Italic := False;
  CurrentCell.Interior.Color := Color;
  CurrentCell.Borders.LineStyle := 1;
}

; Item formula
TempFormula := '=Indirect(address('+sList.Strings[cList.IndexOf(IntToStr(RowId))]+','+IntToStr(ColId)+',,,"Cost"))*Indirect(address('+sList.Strings[bList.IndexOf("-2")]+','+IntToStr(ColId)+',,,"Cost"))';
CurrentCell := HelpSheet.Cells[RowId][ColId];
CurrentCell.Formula := TempFormula;
/*CurrentCell.NumberFormat := CellPriceFormat;*/
CurrentCell.Font.Italic := %IF{@%DB_RES_PRICE%,False,True};
CurrentCell.Interior.Color := Color;
CurrentCell.Borders.LineStyle := 1;



;supplier
s_colid:=colid+1;
currentcell:=costsheet.cells[rowid+row_increase][s_colid];
currentcell.value:="%DSP_COST_SUPPLIER%";
currentcell.borders.linestyle:=1;

;unit
u_colid:=colid-1;
currentcell:=costsheet.cells[rowid+row_increase][u_colid];
u_recent_value:=currentcell.value;
if "@%DB_COST_ASSEMBLY%"="" then
{
	costsheet.cells[rowid+row_increase][u_colid+1].value:=u_recent_value;
	currentcell.value:="";
	/*costsheet.cells[rowid+row_increase][u_colid-1].value:=1;*/
}
else
{
	tot_formula:="="+RId+LBr+inttostr(0)+RBr+CId+Lbr+"1"+RBr+"/@COST_QUANTITY";
	currentcell.formulaR1C1:=tot_formula;
}
currentcell.borders.linestyle:=1;

;quantity per surface
wps_colid:=colid-2;
currentcell:=costsheet.cells[rowid+row_increase][wps_colid];
if "@%DB_COST_ASSEMBLY%"<>"" then
	currentcell.formulaR1C1:="=@COST_QUANTITY/mianji";
else
	currentcell.value:="";
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

CostSheet.Rows[RowId+1+row_increase].select();
excel.Selection.EntireRow.Insert();
excel.Selection.EntireRow.Insert();

tmp_rowid_increase:=RowId+row_increase;

CostSheet.Range[CostSheet.Cells[tmp_rowid_increase+1][1]][CostSheet.Cells[tmp_rowid_increase+2][1]].merge();
CostSheet.Cells[tmp_rowid_increase+1][1].Value:="小计";
costsheet.cells[tmp_rowid_increase+1][1].VerticalAlignment:=-4108;
costsheet.cells[tmp_rowid_increase+1][1].HorizontalAlignment:=-4108;
CostSheet.Range[CostSheet.Cells[tmp_rowid_increase+1][2]][CostSheet.Cells[tmp_rowid_increase+1][3]].merge();
CostSheet.Cells[tmp_rowid_increase+1][2].Value:="附件损耗";
CostSheet.Cells[tmp_rowid_increase+2][2].Value:="附件小计";
CostSheet.Range[CostSheet.Cells[tmp_rowid_increase+2][2]][CostSheet.Cells[tmp_rowid_increase+2][3]].merge();
CostSheet.Range[CostSheet.Cells[tmp_rowid_increase+1][5]][CostSheet.Cells[tmp_rowid_increase+1][7]].merge();

/*CostSheet.Cells[tmp_rowid_increase+1][5].formula:='=Indirect("Cost!"&address('+sList.Strings[cList.IndexOf(IntToStr(RowId))]+","+IntToStr(ColC1)+"))";*/
CostSheet.Cells[tmp_rowid_increase+1][5].value:=0;

CostSheet.Cells[tmp_rowid_increase+1][5].NumberFormatLocal:="0.0%";
CostSheet.Range[CostSheet.Cells[tmp_rowid_increase+2][5]][CostSheet.Cells[tmp_rowid_increase+2][7]].merge();
costsheet.cells[tmp_rowid_increase+2][5].NumberFormat:=CellCostFormat;

Formula1 := "="+SumFormulaText+"("+RId+LBr+IntToStr(recent_rowid-tmp_rowid_increase-2)+RBr+CId+LBr+"2"+RBr+":"+RId+LBr+"-2"+RBr+CId+Lbr+"2"+RBr+")*(1+"+RId+LBr+"-1"+RBr+CId+")";
CostSheet.Cells[tmp_rowid_increase+2][5].FormulaR1C1:=Formula1;


row_increase:=row_increase+2;

/*
CostSheet.Range[costsheet.cells[RowId+1][1]][costsheet.cells[Rowid+2][1]].merge;*/

CostSheet.Range[CostSheet.Cells[tmp_rowid_increase+1][1]][CostSheet.Cells[tmp_rowid_increase+1][8]].Interior.Color:=14935011;
CostSheet.Range[CostSheet.Cells[tmp_rowid_increase+2][1]][CostSheet.Cells[tmp_rowid_increase+2][8]].Interior.Color:=14935011;

RowId_1:=tmp_rowid_increase+2;
