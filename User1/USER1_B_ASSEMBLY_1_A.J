/* USER1_B_ASSEMBLY_1_A.J
 * show accessory prices */



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
if recent_rowid=-1 || recent_rowid>(rowid+row_increase) then
	recent_rowid:=rowid+row_increase;




;quantity per surface
wps_colid:=colid-2;
currentcell:=costsheet.cells[rowid+row_increase][wps_colid];
if "@%DB_ATTRIB_NO%"<>"" then
	/*currentcell.formulaR1C1:="=@COST_QUANTITY/mianji";*/
    currentcell.formulaR1C1:="=@%DB_ATTRIB_NO%/%ASSEMBLYCOUNT%";
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
/*
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


CostSheet.Cells[tmp_rowid_increase+1][5].value:=0;

CostSheet.Cells[tmp_rowid_increase+1][5].NumberFormatLocal:="0.0%";
CostSheet.Range[CostSheet.Cells[tmp_rowid_increase+2][5]][CostSheet.Cells[tmp_rowid_increase+2][7]].merge();
costsheet.cells[tmp_rowid_increase+2][5].NumberFormat:=CellCostFormat;

Formula1 := "="+SumFormulaText+"("+RId+LBr+IntToStr(recent_rowid-tmp_rowid_increase-2)+RBr+CId+LBr+"2"+RBr+":"+RId+LBr+"-2"+RBr+CId+Lbr+"2"+RBr+")*(1+"+RId+LBr+"-1"+RBr+CId+")";
CostSheet.Cells[tmp_rowid_increase+2][5].FormulaR1C1:=Formula1;


row_increase:=row_increase+2;



CostSheet.Range[CostSheet.Cells[tmp_rowid_increase+1][1]][CostSheet.Cells[tmp_rowid_increase+1][8]].Interior.Color:=14935011;
CostSheet.Range[CostSheet.Cells[tmp_rowid_increase+2][1]][CostSheet.Cells[tmp_rowid_increase+2][8]].Interior.Color:=14935011;
*/

