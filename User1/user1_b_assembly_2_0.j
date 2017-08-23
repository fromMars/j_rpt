
; ******************************Estim Excel************************************
; *****************************************************************************
; %NAME% (%BATCH%)
; user1_b_assembly_2_0.j
;


%% detail
; ******************************Estim Excel************************************
; *****************************************************************************
; %NAME% (%BATCH%) - Detail   b_assembly_2.j
; 



; Item price

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


tmp_rowid_increase:=RowId+row_increase;

CostSheet.Cells[tmp_rowid_increase+1][1].Value:="A";
CostSheet.Range[CostSheet.Cells[tmp_rowid_increase+1][2]][CostSheet.Cells[tmp_rowid_increase+1][2]].merge();
CostSheet.Cells[tmp_rowid_increase+1][2].Value:="���Ϸ�С��";

CostSheet.Range[CostSheet.Cells[tmp_rowid_increase+1][3]][CostSheet.Cells[tmp_rowid_increase+1][8]].merge();

s0:=RId+LBr+IntToStr(RowId_0-tmp_rowid_increase-1)+RBr+CId+LBr+"2"+RBr;
s1:=RId+LBr+IntToStr(RowId_1-tmp_rowid_increase-1)+RBr+CId+LBr+"2"+RBr;
s2:=RId+LBr+IntToStr(RowId_2-tmp_rowid_increase-1)+RBr+CId+LBr+"2"+RBr;
if RowId_0=0 then
	s0 := "0";
if RowId_1=0 then
	s1 := "0";
if RowId_2=0 then
	s2 := "0";


Formula0 := "="+SumFormulaText+"("+s0+","+s1+","+s2+")";
CostSheet.Cells[tmp_rowid_increase+1][3].formulaR1C1:=formula0;

/*
CostSheet.Range[CostSheet.Cells[RowId+1][1]][CostSheet.Cells[RowId+1][8]].Interior.Color:=14935011;*/
CostSheet.Range[CostSheet.Cells[tmp_rowid_increase+1][1]][CostSheet.Cells[tmp_rowid_increase+1][8]].Interior.Color:=16777215;


/*a_fee_row:=tmp_rowid_increase+1;*/
RowId_A:=tmp_rowid_increase+1;
row_increase:=row_increase+1;

