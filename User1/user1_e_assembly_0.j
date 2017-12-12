/* USER1_E_ASSEMBLY_0.J
 * 
 *                      */


costsheet.columns[4].delete();
costsheet.range[costsheet.columns[8]][costsheet.columns[16]].delete();

s_index := bList.IndexOf("-3");

CostSheet.Range[costsheet.cells[3][1]][costsheet.cells[rowid+2][7]].Borders.LineStyle:=1;

/* if glass price is not 0, the artikel will appera in the cost sheet, so
 * we need to remove the empty artikel 20 line above the custom glass lines.*/
if glass_price=1 then
    CostSheet.rows[RowId_0+1].entirerow.delete();


if s_index <> -1 then
{
    CostSheet.Columns.Autofit;
}


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

