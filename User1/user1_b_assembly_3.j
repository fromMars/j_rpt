
recent_rowid:=-1;

; ******************************Estim Excel************************************
; *****************************************************************************
; %NAME% (%BATCH%)
; 


%% detail
; ******************************Estim Excel************************************
; *****************************************************************************
; %NAME% (%BATCH%) - Detail   b_assembly_3.j
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
if u_recent_value<>0 then
{
	costsheet.cells[rowid+row_increase][u_colid+1].value:=u_recent_value;
	currentcell.value:="";
	/*costsheet.cells[rowid+row_increase][u_colid-1].value:=1;*/
}
else
{
	tot_formula:="="+RId+LBr+inttostr(0)+RBr+CId+Lbr+"1"+RBr+"/@%DB_COST_FRAME%";
	currentcell.formulaR1C1:=tot_formula;
}
currentcell.borders.linestyle:=1;

;quantity per surface
wps_colid:=colid-2;
currentcell:=costsheet.cells[rowid+row_increase][wps_colid];
if @%DB_COST_FRAME%<>1 then
	currentcell.formulaR1C1:="=@%DB_COST_FRAME%/mianji";
else
	currentcell.formulaR1C1:="=@%DB_COST_FRAME%";
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

/*
CostSheet.Range[CostSheet.Cells[tmp_rowid_increase+1][1]][CostSheet.Cells[tmp_rowid_increase+2][1]].merge();*/
CostSheet.Cells[tmp_rowid_increase+1][1].Value:="B";
CostSheet.Cells[tmp_rowid_increase+2][1].Value:="C";

/*
CostSheet.Range[CostSheet.Cells[tmp_rowid_increase+1][2]][CostSheet.Cells[tmp_rowid_increase+1][3]].merge();*/

CostSheet.Cells[tmp_rowid_increase+1][2].Value:="人工机械直接费";
CostSheet.Cells[tmp_rowid_increase+2][2].Value:="直接费合计";

/*
CostSheet.Range[CostSheet.Cells[tmp_rowid_increase+2][2]][CostSheet.Cells[tmp_rowid_increase+2][3]].merge();*/

CostSheet.Range[CostSheet.Cells[tmp_rowid_increase+1][3]][CostSheet.Cells[tmp_rowid_increase+1][8]].merge();
Formula0 := "="+SumFormulaText+"("+RId+LBr+IntToStr(recent_rowid-tmp_rowid_increase-1)+RBr+CId+LBr+"4"+RBr+":"+RId+LBr+"-1"+RBr+CId+Lbr+"4"+RBr+")";
CostSheet.Cells[tmp_rowid_increase+1][3].formula:=formula0;
/*CostSheet.Cells[tmp_rowid_increase+1][3].NumberFormatLocal:="0.0%";*/

CostSheet.Range[CostSheet.Cells[tmp_rowid_increase+2][3]][CostSheet.Cells[tmp_rowid_increase+2][8]].merge();

/*
Formula1 := "="+SumFormulaText+"("+RId+LBr+IntToStr(recent_rowid-tmp_rowid_increase-2)+RBr+CId+LBr+"2"+RBr+":"+RId+LBr+"-2"+RBr+CId+Lbr+"2"+RBr+")";*/

if a_fee_row=0 then
{
	Formula1 := "="+SumFormulaText+"(0,"+RId+LBr+"-1"+RBr+CId+Lbr+"0"+RBr+")";
}
else
{
	Formula1 := "="+SumFormulaText+"("+RId+LBr+inttostr(tmp_rowid_increase+2-a_fee_row-1)+RBr+Cid+","+RId+LBr+"-1"+RBr+CId+Lbr+"0"+RBr+")";
}
CostSheet.Cells[tmp_rowid_increase+2][3].FormulaR1C1:=Formula1;


row_increase:=row_increase+2;












