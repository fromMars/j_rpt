/* USER1_B_ASSEMBLY_1_C.J
 * get handling tab values from 
 * charge table correct quantity 
 * per assembly */



; ******************************Estim Excel************************************
; *****************************************************************************
; %NAME% (%BATCH%)
; 


%% detail
; ******************************Estim Excel************************************
; *****************************************************************************
; %NAME% (%BATCH%) - Detail   b_assembly_1_C.j
; 



; Item price
RowId  := StrToNum(cList.Strings[bList.IndexOf("@%DB_COST_ARTICLE%"+"@%DB_COST_LOSSTYPE%")]);
if recent_rowid=-1 || recent_rowid>(rowid+row_increase) then
	recent_rowid:=rowid+row_increase;


;quantity per surface
/*单樘用量*/
wps_colid:=colid-2;
currentcell:=costsheet.cells[rowid+row_increase][wps_colid];            /*单樘用量*/
currentcell0:=costsheet.cells[rowid+row_increase][wps_colid-2];         /*单位*/
currentcell0.HorizontalAlignment:=-4108;
if @%DB_CHARGE_QUANTITY%<>0 then
{
    currentcell.formulaR1C1:="=@%DB_CHARGE_QUANTITY%/%ASSEMBLYCOUNT%";
    currentcell0.value:="处";
}    
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


