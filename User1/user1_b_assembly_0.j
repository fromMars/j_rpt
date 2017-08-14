Count:=recent_count;
ColId:=recent_colid;
CostSheet:=recent_cost_sheet;

CostSheet.Copy(CostSheet);
cnt := Template.WorkSheets.Count-3;

template_cost:=Template.WorkSheets[cnt];
template_cost.Name := "Cost_"+trim("%ASSEMBLY_TEXT%");
template_cost.Activate();

CostSheet:=template_cost;

CostSheet.Range["RateRows"].Delete;
ColId:=ColId-8;


recent_rowid:=-1;

; ******************************Estim Excel************************************
; *****************************************************************************
; %NAME% (%BATCH%)  b_assembly_0.j
; 

Count := Count + 1;
ColId := ColId + 1;
Color := DataSheet.Range["CellFormat"].Interior.Color;

; Initialize prices assembly level
i := 0;
while (i < cList.Count) do
{
  RowId       := StrToNum(cList.Strings[i]);
  TempValue   := 0.0;
  CurrentCell := CostSheet.Cells[RowId][ColId];
  CurrentCell.Value := TempValue;
  CurrentCell.NumberFormat := CellPriceFormat;
  CurrentCell.Font.Italic := True;
  CurrentCell.Interior.Color := Color;
  CurrentCell.Borders.LineStyle := 1;
  i := i + 1;
};

%% detail
; ******************************Estim Excel************************************
; *****************************************************************************
; %NAME% (%BATCH%) - Detail
; 

/*
; Item counter
RowId       := 1;
TempValue   := Count;
CurrentCell := CostSheet.Cells[RowId][ColId];
CurrentCell.Value := TempValue;
CurrentCell.NumberFormat := CellLineFormat;
CurrentCell.Interior.Color := Color;
CurrentCell.Borders.LineStyle := 1;*/

/*
; Item description
RowId       := 2;
TempValue   := "@%DB_COST_ID%";
CurrentCell := CostSheet.Cells[RowId][ColId];
CurrentCell.Value := TempValue;
CurrentCell.NumberFormat := CellTextFormat;
CurrentCell.Interior.Color := Color;
CurrentCell.Borders.LineStyle := 1;

; Item width
RowId       := 3;
TempValue   := "";
CurrentCell := CostSheet.Cells[RowId][ColId];
CurrentCell.Value := TempValue;
CurrentCell.NumberFormat := CellTextFormat;
CurrentCell.Interior.Color := Color;
CurrentCell.Borders.LineStyle := 1;

; Item height
RowId       := 4;
TempValue   := "";
CurrentCell := CostSheet.Cells[RowId][ColId];
CurrentCell.Value := TempValue;
CurrentCell.NumberFormat := CellTextFormat;
CurrentCell.Interior.Color := Color;
CurrentCell.Borders.LineStyle := 1;

; Item surface
RowId       := 5;
TempValue   := "";
CurrentCell := CostSheet.Cells[RowId][ColId];
CurrentCell.Value := TempValue;
CurrentCell.NumberFormat := CellTextFormat;
CurrentCell.Interior.Color := Color;
CurrentCell.Borders.LineStyle := 1;
*/


; Item price
/*RowId  := StrToNum(cList.Strings[bList.IndexOf("@%DB_COST_ARTICLE%"+"@%DB_COST_LOSSTYPE%"+"@%DB_COST_RATION%"+"@%DB_COST_FACTOR%"+"@%DB_COST_RATIO%")]);*/
RowId  := StrToNum(cList.Strings[bList.IndexOf("@%DB_PIECE_ARTICLE%"+"@%DB_PIECE_LOSSTYPE%")]);
if recent_rowid=-1 then
	recent_rowid:=rowid;

CellCT := "Indirect(address("+sList.Strings[cList.IndexOf(IntToStr(RowId))]+","+IntToStr(ColCT)+"))/"+pList.Strings[cList.IndexOf(IntToStr(RowId))];
CellC1 := 'Indirect("Cost!"&address('+sList.Strings[cList.IndexOf(IntToStr(RowId))]+","+IntToStr(ColC1)+"))";
CellC2 := 'Indirect("Cost!"&address('+sList.Strings[cList.IndexOf(IntToStr(RowId))]+","+IntToStr(ColC2)+"))";
CellC7 := 'Indirect("Cost!"&address('+sList.Strings[cList.IndexOf(IntToStr(RowId))]+","+IntToStr(ColC7)+"))";
CellC3 := 'Indirect("Cost!"&address('+sList.Strings[cList.IndexOf(IntToStr(RowId))]+","+IntToStr(ColC3)+"))";
CellC4 := 'Indirect("Cost!"&address('+sList.Strings[cList.IndexOf(IntToStr(RowId))]+","+IntToStr(ColC4)+"))";
CellC5 := 'Indirect("Cost!"&address('+sList.Strings[cList.IndexOf(IntToStr(RowId))]+","+IntToStr(ColC5)+"))";
CellC6 := 'Indirect("Cost!"&address('+sList.Strings[cList.IndexOf(IntToStr(RowId))]+","+IntToStr(ColC6)+"))";
if (StrToNum(StrReplace(pList.Strings[cList.IndexOf(IntToStr(RowId))],"%DECIMALSEP%","."),0) > 0) then
{
  TempValue   := StrReplace("@%DB_PIECE_PRICE%/%ASSEMBLYCOUNT%",".","%DECIMALSEP%");
  TempFormula := "=((((((((("+TempValue+")*("+CellCT+"))*(1+"+CellC1+"))*(1-"+CellC2+"))*"+CellC7+")*"+CellC3+")*(1+"+CellC6+"))*(1+"+CellC4+"))*(1-"+CellC5+"))";
  CurrentCell := CostSheet.Cells[RowId][ColId];
  CurrentCell.Formula := TempFormula;
  CurrentCell.NumberFormat := CellPriceFormat;
  CurrentCell.Font.Italic := False;
  CurrentCell.Interior.Color := Color;
  CurrentCell.Borders.LineStyle := 1;
}
else
{
  TempValue   := StrReplace("@%DB_PIECE_PRICE%/%ASSEMBLYCOUNT%",".","%DECIMALSEP%");
  TempFormula := "=(((((((("+TempValue+")*(1+"+CellC1+"))*(1-"+CellC2+"))*"+CellC7+")*"+CellC3+")*(1+"+CellC6+"))*(1+"+CellC4+"))*(1-"+CellC5+"))";
  CurrentCell := CostSheet.Cells[RowId][ColId];
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
currentcell:=costsheet.cells[rowid][s_colid];
currentcell.value:="%DSP_PIECE_SUPPLIER%";
currentcell.borders.linestyle:=1;

;weight
u_colid:=colid-2;
currentcell:=costsheet.cells[rowid][u_colid];
currentcell.value:="@%DB_PIECE_WEIGHT%";
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


CostSheet.Rows[RowId+1].select();
excel.Selection.EntireRow.Insert();
excel.Selection.EntireRow.Insert();
row_increase:=2;

CostSheet.Range[CostSheet.Cells[RowId+1][1]][CostSheet.Cells[RowId+2][1]].merge();
CostSheet.Cells[RowId+1][1].Value:="小计";
CostSheet.Range[CostSheet.Cells[RowId+1][2]][CostSheet.Cells[RowId+1][3]].merge();
CostSheet.Cells[RowId+1][2].Value:="型材损耗";
CostSheet.Range[CostSheet.Cells[RowId+2][2]][CostSheet.Cells[RowId+2][3]].merge();
CostSheet.Cells[RowId+2][2].Value:="型材小计";
CostSheet.Range[CostSheet.Cells[RowId+1][4]][CostSheet.Cells[RowId+1][6]].merge();
CostSheet.Cells[RowId+1][4].formula:="="+CellC1;
CostSheet.Cells[RowId+1][4].NumberFormatLocal:="0.0%";
CostSheet.Range[CostSheet.Cells[RowId+2][4]][CostSheet.Cells[RowId+2][6]].merge();

Formula1 := "="+SumFormulaText+"("+RId+LBr+IntToStr(recent_rowid-rowid-2)+RBr+CId+LBr+"3"+RBr+":"+RId+LBr+"-2"+RBr+CId+Lbr+"3"+RBr+")";
CostSheet.Cells[RowId+2][4].FormulaR1C1:=Formula1;

/*CostSheet.Range[costsheet.cells[RowId+1][1].address+":"+costsheet.cells[Rowid+2][1].address].merge;*/





