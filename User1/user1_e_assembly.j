
CostSheet.Rows[RowId+1+row_increase].select();
excel.Selection.EntireRow.Insert();
excel.Selection.EntireRow.Insert();

tmp_rowid_increase:=RowId+row_increase;

/*
CostSheet.Range[CostSheet.Cells[tmp_rowid_increase+1][1]][CostSheet.Cells[tmp_rowid_increase+2][1]].merge();*/
CostSheet.Cells[tmp_rowid_increase+1][1].Value:="B";
CostSheet.Cells[tmp_rowid_increase+2][1].Value:="C";
CostSheet.Cells[tmp_rowid_increase+1][1].VerticalAlignment:=-4108;
CostSheet.Cells[tmp_rowid_increase+1][1].HorizontalAlignment:=-4108;
CostSheet.Cells[tmp_rowid_increase+2][1].VerticalAlignment:=-4108;
CostSheet.Cells[tmp_rowid_increase+2][1].HorizontalAlignment:=-4108;

/*
CostSheet.Range[CostSheet.Cells[tmp_rowid_increase+1][2]][CostSheet.Cells[tmp_rowid_increase+1][3]].merge();*/

CostSheet.Cells[tmp_rowid_increase+1][2].Value:="人工机械直接费";
CostSheet.Cells[tmp_rowid_increase+2][2].Value:="直接费合计";
CostSheet.Cells[tmp_rowid_increase+2][8].Value:="A+B";

/*
CostSheet.Range[CostSheet.Cells[tmp_rowid_increase+2][2]][CostSheet.Cells[tmp_rowid_increase+2][3]].merge();*/

CostSheet.Range[CostSheet.Cells[tmp_rowid_increase+1][3]][CostSheet.Cells[tmp_rowid_increase+1][7]].merge();
Formula0 := "="+SumFormulaText+"("+RId+LBr+IntToStr(recent_rowid-tmp_rowid_increase-1)+RBr+CId+LBr+"4"+RBr+":"+RId+LBr+"-1"+RBr+CId+Lbr+"4"+RBr+")";
if recent_rowid>RowId_A then
    CostSheet.Cells[tmp_rowid_increase+1][3].formula:=formula0;
else
    costsheet.cells[tmp_rowid_increase+1][3].value:=0;
costsheet.cells[tmp_rowid_increase+1][3].NumberFormat:=CellCostFormat;

/*CostSheet.Cells[tmp_rowid_increase+1][3].NumberFormatLocal:="0.0%";*/

CostSheet.Range[CostSheet.Cells[tmp_rowid_increase+2][3]][CostSheet.Cells[tmp_rowid_increase+2][7]].merge();

/*
Formula1 := "="+SumFormulaText+"("+RId+LBr+IntToStr(recent_rowid-tmp_rowid_increase-2)+RBr+CId+LBr+"2"+RBr+":"+RId+LBr+"-2"+RBr+CId+Lbr+"2"+RBr+")";*/

if RowId_A=0 then
{
	Formula1 := "="+SumFormulaText+"(0,"+RId+LBr+"-1"+RBr+CId+Lbr+"0"+RBr+")";
}
else
{
	Formula1 := "="+SumFormulaText+"("+RId+LBr+inttostr(-(tmp_rowid_increase+2-RowId_A))+RBr+Cid+","+RId+LBr+"-1"+RBr+CId+Lbr+"0"+RBr+")";
}
CostSheet.Cells[tmp_rowid_increase+2][3].FormulaR1C1:=Formula1;
CostSheet.Cells[tmp_rowid_increase+2][3].NumberFormat:=CellCostFormat;

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
CurrentCell.NumberFormat := CellCostFormat;
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

list_no_formula:="=row()-"+inttostr(row_increase+2);
/*add indirect fees, planned vats, vats etc.*/
rowid:=rowid+row_increase;
/*list_no:=rowid-row_increase-3+1;*/
RowId_C:=rowid;
costsheet.cells[rowid+1][1].formula:=list_no_formula;
costsheet.cells[rowid+1][2].value:="现场管理费";
costsheet.range[costsheet.cells[rowid+1][3]][costsheet.cells[rowid+1][6]].merge();
costsheet.cells[rowid+1][3].value:=0;
costsheet.cells[rowid+1][3].NumberFormatLocal:="0.0%";
costsheet.cells[rowid+1][7].FormulaR1C1:="="+RId+LBr+"-1"+RBr+CId+LBr+"-4"+RBr+"*"+RId+CId+LBr+"-4"+RBr;

rowid:=rowid+1;
/*list_no:=list_no+1;*/
costsheet.cells[rowid+1][1].formula:=list_no_formula;
costsheet.cells[rowid+1][2].value:="企业管理费";
costsheet.range[costsheet.cells[rowid+1][3]][costsheet.cells[rowid+1][6]].merge();
costsheet.cells[rowid+1][3].value:=0;
costsheet.cells[rowid+1][3].NumberFormatLocal:="0.0%";
costsheet.cells[rowid+1][7].FormulaR1C1:="="+RId+LBr+"-2"+RBr+CId+LBr+"-4"+RBr+"*"+RId+CId+LBr+"-4"+RBr;

rowid:=rowid+1;
/*list_no:=list_no+1;*/
costsheet.cells[rowid+1][1].value:="D";
costsheet.cells[rowid+1][1].VerticalAlignment:=-4108;
costsheet.cells[rowid+1][1].HorizontalAlignment:=-4108;
costsheet.cells[rowid+1][2].value:="间接费合计";
costsheet.range[costsheet.cells[rowid+1][3]][costsheet.cells[rowid+1][7]].merge();
costsheet.cells[rowid+1][3].value:=0;
costsheet.cells[rowid+1][3].FormulaR1C1:="=sum("+RId+LBr+"-2"+RBr+CId+LBr+"4"+RBr+":"+RId+LBr+"-1"+RBr+CId+LBr+"4"+RBr+")";

rowid:=rowid+1;
/*list_no:=list_no+1;*/
costsheet.cells[rowid+1][1].value:="E";
costsheet.cells[rowid+1][1].VerticalAlignment:=-4108;
costsheet.cells[rowid+1][1].HorizontalAlignment:=-4108;
costsheet.cells[rowid+1][2].value:="小计";
costsheet.range[costsheet.cells[rowid+1][3]][costsheet.cells[rowid+1][7]].merge();
costsheet.cells[rowid+1][3].value:=0;
costsheet.cells[rowid+1][3].NumberFormat:=CellCostFormat;
costsheet.cells[rowid+1][3].FormulaR1C1:="=sum("+RId+LBr+inttostr(RowId_C-rowid-1)+RBr+CId+LBr+"0"+RBr+","+RId+LBr+"-1"+RBr+CId+LBr+"0"+RBr+")";
costsheet.cells[rowid+1][8].value:="C+D";

rowid:=rowid+1;
/*list_no:=list_no+1;*/
costsheet.cells[rowid+1][1].value:="F";
costsheet.cells[rowid+1][1].VerticalAlignment:=-4108;
costsheet.cells[rowid+1][1].HorizontalAlignment:=-4108;
costsheet.cells[rowid+1][2].value:="计划利润";
costsheet.range[costsheet.cells[rowid+1][3]][costsheet.cells[rowid+1][6]].merge();
costsheet.cells[rowid+1][3].value:=0;
costsheet.cells[rowid+1][3].NumberFormatLocal:="0.0%";
costsheet.cells[rowid+1][7].FormulaR1C1:="="+RId+LBr+"-1"+RBr+CId+LBr+"-4"+RBr+"*"+RId+LBr+"0"+RBr+CId+LBr+"-4"+RBr;

rowid:=rowid+1;
/*list_no:=list_no+1;*/
costsheet.cells[rowid+1][1].value:="G";
costsheet.cells[rowid+1][1].VerticalAlignment:=-4108;
costsheet.cells[rowid+1][1].HorizontalAlignment:=-4108;
costsheet.cells[rowid+1][2].value:="税金";
costsheet.range[costsheet.cells[rowid+1][3]][costsheet.cells[rowid+1][6]].merge();
costsheet.cells[rowid+1][3].value:=0;
costsheet.cells[rowid+1][3].NumberFormatLocal:="0.0%";
costsheet.cells[rowid+1][7].FormulaR1C1:="="+RId+LBr+"-2"+RBr+CId+LBr+"-4"+RBr+"*"+RId+LBr+"0"+RBr+CId+LBr+"-4"+RBr;

rowid:=rowid+1;
/*list_no:=list_no+1;*/
costsheet.cells[rowid+1][1].value:="H";
costsheet.cells[rowid+1][1].VerticalAlignment:=-4108;
costsheet.cells[rowid+1][1].HorizontalAlignment:=-4108;
costsheet.cells[rowid+1][2].value:="合计（元/樘）";
costsheet.range[costsheet.cells[rowid+1][3]][costsheet.cells[rowid+1][7]].merge();
costsheet.cells[rowid+1][3].value:=0;
costsheet.cells[rowid+1][3].NumberFormat:=CellCostFormat;
costsheet.cells[rowid+1][3].FormulaR1C1:="=sum("+RId+LBr+"-2"+RBr+CId+LBr+"4"+RBr+","+RId+LBr+"-1"+RBr+CId+LBr+"4"+RBr+","+RId+LBr+"-3"+RBr+CId+LBr+"0"+RBr+")";
costsheet.cells[rowid+1][8].value:="E+F+G";


rowid:=rowid+1;
/*list_no:=list_no+1;*/
costsheet.cells[rowid+1][1].value:="I";
costsheet.cells[rowid+1][1].VerticalAlignment:=-4108;
costsheet.cells[rowid+1][1].HorizontalAlignment:=-4108;
costsheet.cells[rowid+1][2].value:="单价（元/㎡）";
costsheet.range[costsheet.cells[rowid+1][3]][costsheet.cells[rowid+1][7]].merge();
costsheet.cells[rowid+1][3].value:=0;
costsheet.cells[rowid+1][3].NumberFormat:=CellCostFormat;
costsheet.cells[rowid+1][3].FormulaR1C1:="="+RId+LBr+"-1"+RBr+CId+LBr+"0"+RBr+"/mianji";
/*costsheet.cells[rowid+1][8].value:="(E+F+G)/面积";*/

costsheet.range["danjia"].formula:="="+costsheet.cells[rowid+1][3].address;


rowid:=rowid+1;
costsheet.range[costsheet.cells[rowid+1][1]][costsheet.cells[rowid+2][8]].merge();
costsheet.cells[rowid+1][1].value:="                                制单人："+"                                                                "+"批准：";
costsheet.cells[rowid+1][1].VerticalAlignment:=-4108;    /*xlCenter: -4108  xlLeft: -4131  xlRight: -4152*/
/*costsheet.cells[rowid+1][1].Font.Size:=12;*/


