/* USER1_B_ASSEMBLY_1_G.J
 * add custom glass prices */

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
/*插入行*/
rowid:=rowid+1;
costsheet.rows[rowid].select();
excel.selection.entirerow.insert();
row_increase:=row_increase+1;
CostSheet.Range[CostSheet.Cells[RowId][1]][CostSheet.Cells[RowId][8]].Interior.Color:=16777215;             /*白色*/

/*输出玻璃信息*/
costsheet.cells[rowid][1].formula:=list_no_formula;
costsheet.cells[rowid][2].value:="@%DB_GLASS_DESC%";
costsheet.cells[rowid][3].value:="㎡";
costsheet.cells[rowid][3].HorizontalAlignment:=-4108;

;玻璃价格
TempValue   := StrReplace("@%DB_GLASS_PRICE%/%ASSEMBLYCOUNT%",".","%DECIMALSEP%");
if @%DB_GLASS_PRICE%<>0 && RowId_G<>0 then
{
    CellCT := 'Indirect("Cost!"&address('+inttostr(strtonum(sList.Strings[cList.IndexOf(IntToStr(RowId_G))]))+","+IntToStr(ColCT)+"))/"+pList.Strings[cList.IndexOf(IntToStr(RowId_G))];
    CellC1 := 'Indirect("Cost!"&address('+inttostr(strtonum(sList.Strings[cList.IndexOf(IntToStr(RowId_G))]))+","+IntToStr(ColC1)+"))";
    CellC2 := 'Indirect("Cost!"&address('+inttostr(strtonum(sList.Strings[cList.IndexOf(IntToStr(RowId_G))]))+","+IntToStr(ColC2)+"))";
    CellC7 := 'Indirect("Cost!"&address('+inttostr(strtonum(sList.Strings[cList.IndexOf(IntToStr(RowId_G))]))+","+IntToStr(ColC7)+"))";
    CellC3 := 'Indirect("Cost!"&address('+inttostr(strtonum(sList.Strings[cList.IndexOf(IntToStr(RowId_G))]))+","+IntToStr(ColC3)+"))";
    CellC4 := 'Indirect("Cost!"&address('+inttostr(strtonum(sList.Strings[cList.IndexOf(IntToStr(RowId_G))]))+","+IntToStr(ColC4)+"))";
    CellC5 := 'Indirect("Cost!"&address('+inttostr(strtonum(sList.Strings[cList.IndexOf(IntToStr(RowId_G))]))+","+IntToStr(ColC5)+"))";
    CellC6 := 'Indirect("Cost!"&address('+inttostr(strtonum(sList.Strings[cList.IndexOf(IntToStr(RowId_G))]))+","+IntToStr(ColC6)+"))";


    if (StrToNum(StrReplace(pList.Strings[cList.IndexOf(IntToStr(RowId_G))],"%DECIMALSEP%","."),0) > 0) then
    {
        TempFormula := "=((((((((("+TempValue+")*("+CellCT+"))*(1+"+CellC1+"))*(1-"+CellC2+"))*"+CellC7+")*"+CellC3+")*(1+"+CellC6+"))*(1+"+CellC4+"))*(1-"+CellC5+"))";
    }
    else
    {
        TempFormula := "=((((((((("+TempValue+")*("+numtostr(1)+"))*(1+"+CellC1+"))*(1-"+CellC2+"))*"+CellC7+")*"+CellC3+")*(1+"+CellC6+"))*(1+"+CellC4+"))*(1-"+CellC5+"))";
    }
}
else
{
    TempFormula:="="+TempValue;
}

CurrentCell := CostSheet.Cells[rowid][ColId];
CurrentCell.Formula := TempFormula;
CurrentCell.Font.Italic := False;
CurrentCell.Interior.Color := Color;
CurrentCell.Borders.LineStyle := 1;



;供应商
s_colid:=colid+1;
currentcell:=costsheet.cells[rowid][s_colid];
currentcell.value:="%DSP_GLASS_SUPPLIER%";
currentcell.borders.linestyle:=1;

;单价
u_colid:=colid-1;
currentcell:=costsheet.cells[rowid][u_colid];
u_recent_value:=0;
if "@%DB_GLASS_ASSEMBLY%"="" then
{
	costsheet.cells[rowid][u_colid+1].value:=u_recent_value;
	currentcell.value:="";
}
else
{
	tot_formula:="="+RId+LBr+inttostr(0)+RBr+CId+Lbr+"1"+RBr+"/@%DB_GLASS_NO%";
	currentcell.formulaR1C1:=tot_formula;
}
currentcell.borders.linestyle:=1;

;单樘用量
wps_colid:=colid-2;
currentcell:=costsheet.cells[rowid][wps_colid];
if "@%DB_GLASS_ASSEMBLY%"<>"" then
    currentcell.formulaR1C1:="=@%DB_GLASS_SURFACE%/%ASSEMBLYCOUNT%";
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
list_no_formula:="=row()-"+inttostr(row_increase+3);

