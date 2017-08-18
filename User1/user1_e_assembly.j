; ******************************Estim Excel************************************
; *****************************************************************************
; %NAME% (%BATCH%)
; 
/*
ColId:= ColId + 1;
Color := DataSheet.Range["HeadFormat"].Interior.Color;

RowId       := 1;
TempValue   := "";
CurrentCell := CostSheet.Cells[RowId][ColId];
CurrentCell.Value := TempValue;
CurrentCell.NumberFormat := CellTextFormat;
CurrentCell.Interior.Color := Color;
CurrentCell.Borders.LineStyle := 1;
ColumnMerge := CostSheet.Range[CostSheet.Cells[2][ColId]][CostSheet.Cells[5][ColId]];
ColumnMerge.MergeCells := True;
ColumnMerge.Value := DataSheet.Range["ProjectPrice"].Value;
ColumnMerge.NumberFormat := CellTextFormat;
ColumnMerge.Font.Bold := True;
ColumnMerge.Interior.Color := Color;
ColumnMerge.Borders.LineStyle := 1;*/

%% detail
; ******************************Estim Excel************************************
; *****************************************************************************
; %NAME% (%BATCH%) - Detail
; 
/*
; Total item price
search_index := bList.IndexOf("@%DB_COST_ARTICLE%"+"@%DB_COST_LOSSTYPE%"+"@%DB_COST_RATION%"+"@%DB_COST_FACTOR%"+"@%DB_COST_RATIO%");
if search_index >=0 then
{
RowId       := StrToNum(cList.Strings[bList.IndexOf("@%DB_COST_ARTICLE%"+"@%DB_COST_LOSSTYPE%"+"@%DB_COST_RATION%"+"@%DB_COST_FACTOR%"+"@%DB_COST_RATIO%")]);
TempFormula := "="+SumFormulaText+"(Help!"+RId+CId+LBr+IntToStr(-Count)+RBr+":Help!"+RId+CId+LBr+"-1"+RBr+")";
CurrentCell := CostSheet.Cells[RowId][ColId];
CurrentCell.Formula := TempFormula;
CurrentCell.NumberFormat := CellPriceFormat;
CurrentCell.Font.Bold := True;
CurrentCell.Interior.Color := Color;
CurrentCell.Borders.LineStyle := 1;
};
*/
%% break header
; ******************************Estim Excel************************************
; *****************************************************************************
; %NAME% (%BATCH%) - Break header
; 

%% break footer
; ******************************Estim Excel************************************
; *****************************************************************************
; %NAME% (%BATCH%) _ Break footer
; 

%% detail footer
; ******************************Estim Excel************************************
; *****************************************************************************
; %NAME% (%BATCH%) - Detail footer
; 

; Total batch/project price



s_index := bList.IndexOf("-3");

CostSheet.Usedrange.Borders.LineStyle:=1;

if s_index <> -1 then
{
CostSheet.Columns.Autofit;
}

costsheet.columns[4].delete();
costsheet.columns[8].delete();