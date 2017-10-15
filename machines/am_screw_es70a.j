/* ES70A Mitre */

Product     := Param.Product;
Profile     := Param.Profile;
Position    := Param.Workoffset;
Direction   := Param.B;             /* 0 = inside, 1 = outside */
ProfileType := Param.C;             /* 0 = vent, 1 = frame, 2 = T-mullion */

dx:=0;
p_height:=0;
if Product="1030211" then
{
 dx:=4.5;
 p_height:=70;
}
else if Product="1030222" then
{
 dx:=22;
 p_height:=79;
}
/*7.5mm*/
Tool.Name := 'STD_HOLE';
Tool.Profilav:=False;
Tool.X:=Position+dx;
if Product="1030211" then
 Tool.Y:=11;
else if Product="1030222" then
 Tool.Y:=20;
else
 msgbox("1030211 or 1030222 not found!"); 
Tool.Z:=2.5;
Tool.P1X:=0;
Tool.P1Y:=0;
Tool.P1Z:=70;
Tool.P2X:=100;
Tool.P2Y:=0;
Tool.P2Z:=70;
Tool.P3X:=100;
Tool.P3Y:=0;
Tool.P3Z:=0;
Tool.Param['PAR1']:=dx-0.5;
Tool.Param['PAR2']:=7.6;
Tool.Tool:=2;
Tool.Apply();
