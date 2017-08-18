
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



costsheet.columns[4].delete();
costsheet.columns[8].delete();

s_index := bList.IndexOf("-3");

CostSheet.Usedrange.Borders.LineStyle:=1;


if s_index <> -1 then
{
CostSheet.Columns.Autofit;
}
