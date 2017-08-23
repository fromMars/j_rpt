
costsheet.columns[4].delete();
costsheet.columns[8].delete();

s_index := bList.IndexOf("-3");

CostSheet.Usedrange.Borders.LineStyle:=1;


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

; Total batch/project price


