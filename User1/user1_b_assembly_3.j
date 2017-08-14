
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

CellCT := 'Indirect(Cost!address('+numtostr(strtonum(sList.Strings[cList.IndexOf(IntToStr(RowId))])+row_increase)+","+IntToStr(ColCT)+"))/"+pList.Strings[cList.IndexOf(IntToStr(RowId))];
CellC1 := 'Indirect("Cost!"&address('+numtostr(strtonum(sList.Strings[cList.IndexOf(IntToStr(RowId))])+row_increase)+","+IntToStr(ColC1)+"))";
CellC2 := 'Indirect("Cost!"&address('+numtostr(strtonum(sList.Strings[cList.IndexOf(IntToStr(RowId))])+row_increase)+","+IntToStr(ColC2)+"))";
CellC7 := 'Indirect("Cost!"&address('+numtostr(strtonum(sList.Strings[cList.IndexOf(IntToStr(RowId))])+row_increase)+","+IntToStr(ColC7)+"))";
CellC3 := 'Indirect("Cost!"&address('+numtostr(strtonum(sList.Strings[cList.IndexOf(IntToStr(RowId))])+row_increase)+","+IntToStr(ColC3)+"))";
CellC4 := 'Indirect("Cost!"&address('+numtostr(strtonum(sList.Strings[cList.IndexOf(IntToStr(RowId))])+row_increase)+","+IntToStr(ColC4)+"))";
CellC5 := 'Indirect("Cost!"&address('+numtostr(strtonum(sList.Strings[cList.IndexOf(IntToStr(RowId))])+row_increase)+","+IntToStr(ColC5)+"))";
CellC6 := 'Indirect("Cost!"&address('+numtostr(strtonum(sList.Strings[cList.IndexOf(IntToStr(RowId))])+row_increase)+","+IntToStr(ColC6)+"))";
if (StrToNum(StrReplace(pList.Strings[cList.IndexOf(IntToStr(RowId))],"%DECIMALSEP%","."),0) > 0) then
{
  TempValue   := StrReplace("@%DB_RES_PRICE%/%ASSEMBLYCOUNT%",".","%DECIMALSEP%");
  TempFormula := "=((((((((("+TempValue+")*("+CellCT+"))*(1+"+CellC1+"))*(1-"+CellC2+"))*"+CellC7+")*"+CellC3+")*(1+"+CellC6+"))*(1+"+CellC4+"))*(1-"+CellC5+"))";
  CurrentCell := CostSheet.Cells[RowId+row_increase][ColId];
  CurrentCell.Formula := TempFormula;
  CurrentCell.NumberFormat := CellPriceFormat;
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
  CurrentCell.NumberFormat := CellPriceFormat;
  CurrentCell.Font.Italic := False;
  CurrentCell.Interior.Color := Color;
  CurrentCell.Borders.LineStyle := 1;
}

; Item formula
TempFormula := '=Indirect(address('+sList.Strings[cList.IndexOf(IntToStr(RowId))]+','+IntToStr(ColId)+',,,"Cost"))*Indirect(address('+sList.Strings[bList.IndexOf("-2")]+','+IntToStr(ColId)+',,,"Cost"))';
CurrentCell := HelpSheet.Cells[RowId][ColId];
CurrentCell.Formula := TempFormula;
CurrentCell.NumberFormat := CellPriceFormat;
CurrentCell.Font.Italic := %IF{@%DB_RES_PRICE%,False,True};
CurrentCell.Interior.Color := Color;
CurrentCell.Borders.LineStyle := 1;



;supplier
s_colid:=colid+1;
currentcell:=costsheet.cells[rowid+row_increase][s_colid];
currentcell.value:="%DSP_COST_SUPPLIER%";
currentcell.borders.linestyle:=1;

;unit
u_colid:=colid-2;
currentcell:=costsheet.cells[rowid+row_increase][u_colid];
currentcell.value:="@%DB_COST_FRAME%";
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
CostSheet.Range[CostSheet.Cells[tmp_rowid_increase+1][2]][CostSheet.Cells[tmp_rowid_increase+1][3]].merge();
CostSheet.Cells[tmp_rowid_increase+1][2].Value:="人工机械直接费";
CostSheet.Range[CostSheet.Cells[tmp_rowid_increase+2][2]][CostSheet.Cells[tmp_rowid_increase+2][3]].merge();
CostSheet.Cells[tmp_rowid_increase+2][2].Value:="直接费合计";
CostSheet.Range[CostSheet.Cells[tmp_rowid_increase+1][4]][CostSheet.Cells[tmp_rowid_increase+1][6]].merge();
CostSheet.Cells[tmp_rowid_increase+1][4].formula:="="+CellC1;
CostSheet.Cells[tmp_rowid_increase+1][4].NumberFormatLocal:="0.0%";
CostSheet.Range[CostSheet.Cells[tmp_rowid_increase+2][4]][CostSheet.Cells[tmp_rowid_increase+2][6]].merge();
Formula1 := "="+SumFormulaText+"("+RId+LBr+IntToStr(recent_rowid-tmp_rowid_increase-2)+RBr+CId+LBr+"3"+RBr+":"+RId+LBr+"-2"+RBr+CId+Lbr+"3"+RBr+")";
CostSheet.Cells[tmp_rowid_increase+2][4].FormulaR1C1:=Formula1;


row_increase:=row_increase+2;




; Unit price item
RowId       := StrToNum(cList.Strings[bList.IndexOf("-1")]);
/*TempFormula := "="+SumFormulaText+"("+RId+CId+LBr+IntToStr(-Range)+RBr+":"+RId+CId+LBr+"-1"+RBr+")";*/
TempFormula := "="+SumFormulaText+"("+RId+LBr+IntToStr(-Range)+RBr+CId+":"+RId+LBr+"-1"+RBr+CId+")";
CurrentCell := CostSheet.Cells[RowId+row_increase][ColId];
CurrentCell.FormulaR1C1 := TempFormula;
CurrentCell.NumberFormat := CellPriceFormat;
CurrentCell.Font.Italic := False;
CurrentCell.Interior.Color := Color;
CurrentCell.Borders.LineStyle := 1;

; Number of items
RowId       := StrToNum(cList.Strings[bList.IndexOf("-2")]);
TempValue   := %ASSEMBLYCOUNT%;
CurrentCell := CostSheet.Cells[RowId+row_increase][ColId];
CurrentCell.Value := TempValue;
CurrentCell.NumberFormat := CellCountFormat;
CurrentCell.Font.Italic := False;
CurrentCell.Interior.Color := Color;
CurrentCell.Borders.LineStyle := 1;

; Total price item
RowId       := StrToNum(cList.Strings[bList.IndexOf("-3")]);
/*TempFormula := "="+SumFormulaText+"("+RId+CId+LBr+"-2"+RBr+":"+RId+CId+LBr+"-2"+RBr+") seems not multiplied by quantity*/

TempFormula := "="+SumFormulaText+"("+RId+LBr+"-2"+RBr+CId+":"+RId+LBr+"-2"+RBr+CId+")*"+SumFormulaText+"("+RId+LBr+"-1"+RBr+CId+":"+RId+LBr+"-1"+RBr+CId+")";
CurrentCell := CostSheet.Cells[RowId+row_increase][ColId];
CurrentCell.FormulaR1C1 := TempFormula;
CurrentCell.NumberFormat := CellPriceFormat;
CurrentCell.Font.Italic := False;
CurrentCell.Interior.Color := Color;
CurrentCell.Borders.LineStyle := 1;








