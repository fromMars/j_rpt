/* USER1_E_ASSEMBLY.J
 * 
 *                      */


CostSheet.Rows[RowId+1+row_increase].select();
excel.Selection.EntireRow.Insert();
excel.Selection.EntireRow.Insert();

tmp_rowid_increase:=RowId+row_increase;

CostSheet.Cells[tmp_rowid_increase+1][1].Value:="B";
CostSheet.Cells[tmp_rowid_increase+2][1].Value:="C";
CostSheet.Cells[tmp_rowid_increase+1][1].VerticalAlignment:=-4108;
CostSheet.Cells[tmp_rowid_increase+1][1].HorizontalAlignment:=-4108;
CostSheet.Cells[tmp_rowid_increase+2][1].VerticalAlignment:=-4108;
CostSheet.Cells[tmp_rowid_increase+2][1].HorizontalAlignment:=-4108;

CostSheet.Cells[tmp_rowid_increase+1][2].Value:="人工机械直接费";
CostSheet.Cells[tmp_rowid_increase+2][2].Value:="直接费合计";
CostSheet.Cells[tmp_rowid_increase+2][8].Value:="A+B";

CostSheet.Range[CostSheet.Cells[tmp_rowid_increase+1][3]][CostSheet.Cells[tmp_rowid_increase+1][7]].merge();
Formula0 := "="+SumFormulaText+"("+RId+LBr+IntToStr(recent_rowid-tmp_rowid_increase-1)+RBr+CId+LBr+"4"+RBr+":"+RId+LBr+"-1"+RBr+CId+Lbr+"4"+RBr+")";
if recent_rowid>RowId_A then
    CostSheet.Cells[tmp_rowid_increase+1][3].formula:=formula0;
else
    costsheet.cells[tmp_rowid_increase+1][3].value:=0;
costsheet.cells[tmp_rowid_increase+1][3].NumberFormat:=CellCostFormat;

CostSheet.Range[CostSheet.Cells[tmp_rowid_increase+2][3]][CostSheet.Cells[tmp_rowid_increase+2][7]].merge();


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
; e_assembly


%% detail
; ******************************Estim Excel************************************
; *****************************************************************************
; %NAME% (%BATCH%) - Detail
; 

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
rowid:=rowid+row_increase;
RowId_C:=rowid;
costsheet.cells[rowid+1][1].formula:=list_no_formula;
costsheet.cells[rowid+1][2].value:="现场管理费";
costsheet.range[costsheet.cells[rowid+1][3]][costsheet.cells[rowid+1][6]].merge();
costsheet.cells[rowid+1][3].formula:="=Data!CMRate";
costsheet.cells[rowid+1][3].NumberFormatLocal:="0.0%";
costsheet.cells[rowid+1][7].FormulaR1C1:="="+RId+LBr+"-1"+RBr+CId+LBr+"-4"+RBr+"*"+RId+CId+LBr+"-4"+RBr;

rowid:=rowid+1;
costsheet.cells[rowid+1][1].formula:=list_no_formula;
costsheet.cells[rowid+1][2].value:="企业管理费";
costsheet.range[costsheet.cells[rowid+1][3]][costsheet.cells[rowid+1][6]].merge();
costsheet.cells[rowid+1][3].formula:="=Data!SMRate";
costsheet.cells[rowid+1][3].NumberFormatLocal:="0.0%";
costsheet.cells[rowid+1][7].FormulaR1C1:="="+RId+LBr+"-2"+RBr+CId+LBr+"-4"+RBr+"*"+RId+CId+LBr+"-4"+RBr;

rowid:=rowid+1;
costsheet.cells[rowid+1][1].value:="D";
costsheet.cells[rowid+1][1].VerticalAlignment:=-4108;
costsheet.cells[rowid+1][1].HorizontalAlignment:=-4108;
costsheet.cells[rowid+1][2].value:="间接费合计";
costsheet.range[costsheet.cells[rowid+1][3]][costsheet.cells[rowid+1][7]].merge();
costsheet.cells[rowid+1][3].value:=0;
costsheet.cells[rowid+1][3].FormulaR1C1:="=sum("+RId+LBr+"-2"+RBr+CId+LBr+"4"+RBr+":"+RId+LBr+"-1"+RBr+CId+LBr+"4"+RBr+")";

rowid:=rowid+1;
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
costsheet.cells[rowid+1][1].value:="F";
costsheet.cells[rowid+1][1].VerticalAlignment:=-4108;
costsheet.cells[rowid+1][1].HorizontalAlignment:=-4108;
costsheet.cells[rowid+1][2].value:="计划利润";
costsheet.range[costsheet.cells[rowid+1][3]][costsheet.cells[rowid+1][6]].merge();
costsheet.cells[rowid+1][3].formula:="=Data!PPRate";
costsheet.cells[rowid+1][3].NumberFormatLocal:="0.0%";
costsheet.cells[rowid+1][7].FormulaR1C1:="="+RId+LBr+"-1"+RBr+CId+LBr+"-4"+RBr+"*"+RId+LBr+"0"+RBr+CId+LBr+"-4"+RBr;

rowid:=rowid+1;
costsheet.cells[rowid+1][1].value:="G";
costsheet.cells[rowid+1][1].VerticalAlignment:=-4108;
costsheet.cells[rowid+1][1].HorizontalAlignment:=-4108;
costsheet.cells[rowid+1][2].value:="税金";
costsheet.range[costsheet.cells[rowid+1][3]][costsheet.cells[rowid+1][6]].merge();
costsheet.cells[rowid+1][3].formula:="=Data!VRate";
costsheet.cells[rowid+1][3].NumberFormatLocal:="0.0%";
costsheet.cells[rowid+1][7].FormulaR1C1:="=("+RId+LBr+"-2"+RBr+CId+LBr+"-4"+RBr+"+"+RId+LBr+"-1"+RBr+CId+LBr+"0"+RBr+")*"+RId+LBr+"0"+RBr+CId+LBr+"-4"+RBr;

rowid:=rowid+1;
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
costsheet.cells[rowid+1][1].value:="I";
costsheet.cells[rowid+1][1].VerticalAlignment:=-4108;
costsheet.cells[rowid+1][1].HorizontalAlignment:=-4108;
costsheet.cells[rowid+1][2].value:="单价（元/㎡）";
costsheet.range[costsheet.cells[rowid+1][3]][costsheet.cells[rowid+1][7]].merge();
costsheet.cells[rowid+1][3].value:=0;
costsheet.cells[rowid+1][3].NumberFormat:=CellCostFormat;
costsheet.cells[rowid+1][3].FormulaR1C1:="="+RId+LBr+"-1"+RBr+CId+LBr+"0"+RBr+"/mianji";

costsheet.range["danjia"].formula:="="+costsheet.cells[rowid+1][3].address;


rowid:=rowid+1;
costsheet.range[costsheet.cells[rowid+1][1]][costsheet.cells[rowid+2][8]].merge();
costsheet.cells[rowid+1][1].value:="                                制单人："+"                                                                "+"批准：";
/*xlCenter: -4108  xlLeft: -4131  xlRight: -4152*/
costsheet.cells[rowid+1][1].VerticalAlignment:=-4108;


