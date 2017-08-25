
list_no:=rowid-row_increase-3+1;

if RowId_1<>0 then
    rowid:=RowId_1;


; ******************************Estim Excel************************************
; *****************************************************************************
; %NAME% (%BATCH%)
; 


%% detail
; ******************************Estim Excel************************************
; *****************************************************************************
; %NAME% (%BATCH%) - Detail   b_assembly_1_G.j
; 
rowid:=rowid+1;
costsheet.rows[rowid].select();
excel.selection.entirerow.insert();
row_increase:=row_increase+1;
CostSheet.Range[CostSheet.Cells[RowId][1]][CostSheet.Cells[RowId][8]].Interior.Color:=16777215;

/*
if recent_rowid=-1 then
	recent_rowid:=rowid;
*/

costsheet.cells[rowid][1].value:=list_no;
costsheet.cells[rowid][2].value:="@%DB_GLASS_DESC%";

;supplier
s_colid:=colid+1;
currentcell:=costsheet.cells[rowid][s_colid];
currentcell.value:="%DSP_GLASS_SUPPLIER%";
currentcell.borders.linestyle:=1;

;unit
u_colid:=colid-1;
currentcell:=costsheet.cells[rowid][u_colid];
u_recent_value:=0;
if "@%DB_GLASS_ASSEMBLY%"="" then
{
	costsheet.cells[rowid][u_colid+1].value:=u_recent_value;
	currentcell.value:="";
	/*costsheet.cells[rowid+row_increase][u_colid-1].value:=1;*/
}
else
{
	tot_formula:="="+RId+LBr+inttostr(0)+RBr+CId+Lbr+"1"+RBr+"/@%DB_GLASS_NO%";
	currentcell.formulaR1C1:=tot_formula;
}
currentcell.borders.linestyle:=1;

;quantity per surface
wps_colid:=colid-2;
currentcell:=costsheet.cells[rowid][wps_colid];
if "@%DB_GLASS_ASSEMBLY%"<>"" then
	currentcell.formulaR1C1:="=@%DB_GLASS_SURFACE%/mianji";
else
	currentcell.value:="";

currentcell.borders.linestyle:=1;
list_no:=list_no+1;


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

tmp_rowid_increase:=RowId;


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

CostSheet.Range[CostSheet.Cells[RowId+1][1]][CostSheet.Cells[RowId+1][8]].Interior.Color:=14935011;
CostSheet.Range[CostSheet.Cells[RowId+2][1]][CostSheet.Cells[RowId+2][8]].Interior.Color:=14935011;


rowid:=rowid-row_increase;

row_increase:=row_increase+2;
RowId_1:=tmp_rowid_increase+2;

/*
a_fee_row:=tmp_rowid_increase+1;
row_increase:=row_increase+1;
*/
