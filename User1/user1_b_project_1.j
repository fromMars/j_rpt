; ******************************Estim Excel************************************
; *****************************************************************************
; %NAME% (%BATCH%)
; 

Count := Count + 1;
ColId := ColId + 12;
Color := DataSheet.Range["CellFormat"].Interior.Color;

; Initialize prices project level
i := 0;
while (i < cList.Count) do
{
  RowId       := StrToNum(cList.Strings[i]);
  TempValue   := 0.0;
  CurrentCell := CostSheet.Cells[RowId][ColId];
  CurrentCell.Value := TempValue;
  /*CurrentCell.NumberFormat := CellPriceFormat;*/
  CurrentCell.Font.Italic := True;
  /*CurrentCell.Interior.Color := Color;*/
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
RowId  := StrToNum(cList.Strings[bList.IndexOf("@%DB_COST_ARTICLE%"+"@%DB_COST_LOSSTYPE%")]);

/*CellCT := 'Indirect(address('+sList.Strings[cList.IndexOf(IntToStr(RowId))]+","+IntToStr(ColCT)+"))/"+pList.Strings[cList.IndexOf(IntToStr(RowId))];
CellC1 := "Indirect(address("+sList.Strings[cList.IndexOf(IntToStr(RowId))]+","+IntToStr(ColC1)+"))";
CellC2 := "Indirect(address("+sList.Strings[cList.IndexOf(IntToStr(RowId))]+","+IntToStr(ColC2)+"))";
CellC7 := "Indirect(address("+sList.Strings[cList.IndexOf(IntToStr(RowId))]+","+IntToStr(ColC7)+"))";
CellC3 := "Indirect(address("+sList.Strings[cList.IndexOf(IntToStr(RowId))]+","+IntToStr(ColC3)+"))";
CellC4 := "Indirect(address("+sList.Strings[cList.IndexOf(IntToStr(RowId))]+","+IntToStr(ColC4)+"))";
CellC5 := "Indirect(address("+sList.Strings[cList.IndexOf(IntToStr(RowId))]+","+IntToStr(ColC5)+"))";
CellC6 := "Indirect(address("+sList.Strings[cList.IndexOf(IntToStr(RowId))]+","+IntToStr(ColC6)+"))";*/


CellCT := 'Indirect(Cost!address('+sList.Strings[cList.IndexOf(IntToStr(RowId))]+","+IntToStr(ColCT)+"))/"+pList.Strings[cList.IndexOf(IntToStr(RowId))];
CellC1 := 'Indirect("Cost!"&address('+sList.Strings[cList.IndexOf(IntToStr(RowId))]+","+IntToStr(ColC1)+"))";
CellC2 := 'Indirect("Cost!"&address('+sList.Strings[cList.IndexOf(IntToStr(RowId))]+","+IntToStr(ColC2)+"))";
CellC7 := 'Indirect("Cost!"&address('+sList.Strings[cList.IndexOf(IntToStr(RowId))]+","+IntToStr(ColC7)+"))";
CellC3 := 'Indirect("Cost!"&address('+sList.Strings[cList.IndexOf(IntToStr(RowId))]+","+IntToStr(ColC3)+"))";
CellC4 := 'Indirect("Cost!"&address('+sList.Strings[cList.IndexOf(IntToStr(RowId))]+","+IntToStr(ColC4)+"))";
CellC5 := 'Indirect("Cost!"&address('+sList.Strings[cList.IndexOf(IntToStr(RowId))]+","+IntToStr(ColC5)+"))";
CellC6 := 'Indirect("Cost!"&address('+sList.Strings[cList.IndexOf(IntToStr(RowId))]+","+IntToStr(ColC6)+"))";



if (StrToNum(StrReplace(pList.Strings[cList.IndexOf(IntToStr(RowId))],"%DECIMALSEP%","."),0) > 0) then
{
  TempValue   := StrReplace("@%DB_RES_PRICE%",".","%DECIMALSEP%");
  TempFormula := "=((((((((("+TempValue+")*("+CellCT+"))*(1+"+CellC1+"))*(1-"+CellC2+"))*"+CellC7+")*"+CellC3+")*(1+"+CellC6+"))*(1+"+CellC4+"))*(1-"+CellC5+"))";
  CurrentCell := CostSheet.Cells[RowId][ColId];
  CurrentCell.Formula := TempFormula;
/*  CurrentCell.NumberFormat := CellPriceFormat;*/
  CurrentCell.Font.Italic := False;
  /*CurrentCell.Interior.Color := Color;*/
  CurrentCell.Borders.LineStyle := 1;
}
else
{
  TempValue   := StrReplace("@%DB_RES_PRICE%",".","%DECIMALSEP%");
  TempFormula := "=(((((((("+TempValue+")*(1+"+CellC1+"))*(1-"+CellC2+"))*"+CellC7+")*"+CellC3+")*(1+"+CellC6+"))*(1+"+CellC4+"))*(1-"+CellC5+"))";
  CurrentCell := CostSheet.Cells[RowId][ColId];
  CurrentCell.Formula := TempFormula;
/*  CurrentCell.NumberFormat := CellPriceFormat;*/
  CurrentCell.Font.Italic := False;
  /*CurrentCell.Interior.Color := Color;*/
  CurrentCell.Borders.LineStyle := 1;
};

; Item formula
TempFormula := '=Indirect(address('+sList.Strings[cList.IndexOf(IntToStr(RowId))]+','+IntToStr(ColId)+',,,"Cost"))';
CurrentCell := HelpSheet.Cells[RowId][ColId];
CurrentCell.Formula := TempFormula;
/*CurrentCell.NumberFormat := CellPriceFormat;*/
CurrentCell.Font.Italic := %IF{@%DB_RES_PRICE%,False,True};
/*CurrentCell.Interior.Color := Color;*/
CurrentCell.Borders.LineStyle := 1;


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

; Unit price item
RowId       := StrToNum(cList.Strings[bList.IndexOf("-1")]);
/*TempFormula := "="+SumFormulaText+"("+RId+CId++LBr+IntToStr(-Range)+RBr":"+RId+CId+LBr+"-1"+RBr+")";*/
TempFormula := "="+SumFormulaText+"("+RId+LBr+IntToStr(-Range)+RBr+CId+":"+RId+LBr+"-1"+RBr+CId+")";
CurrentCell := CostSheet.Cells[RowId][ColId];
CurrentCell.FormulaR1C1 := TempFormula;
CurrentCell.NumberFormat := CellPriceFormat;
CurrentCell.Font.Italic := False;
/*CurrentCell.Interior.Color := Color;*/
CurrentCell.Borders.LineStyle := 1;

; Number of items
RowId       := StrToNum(cList.Strings[bList.IndexOf("-2")]);
TempValue   := "";
CurrentCell := CostSheet.Cells[RowId][ColId];
CurrentCell.Value := TempValue;
CurrentCell.NumberFormat := CellTextFormat;
CurrentCell.Font.Italic := False;
/*CurrentCell.Interior.Color := Color;*/
CurrentCell.Borders.LineStyle := 1;

; Total price item
RowId       := StrToNum(cList.Strings[bList.IndexOf("-3")]);
/*TempFormula := "="+SumFormulaText+"("+RId+CId+LBr+"-2"+RBr+":"+RId+CId+LBr+"-2"+RBr+")";*/
TempFormula := "="+SumFormulaText+"("+RId+LBr+"-2"+RBr+CId+":"+RId+LBr+"-2"+RBr+CId+")";
CurrentCell := CostSheet.Cells[RowId][ColId];
CurrentCell.FormulaR1C1 := TempFormula;
CurrentCell.NumberFormat := CellPriceFormat;
CurrentCell.Font.Italic := False;
/*CurrentCell.Interior.Color := Color;*/
CurrentCell.Borders.LineStyle := 1;



recent_count:=count;
recent_colid:=colid;
recent_cost_sheet:=CostSheet;