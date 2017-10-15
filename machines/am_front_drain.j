/* DRAIN FRONT */

Product     := Param.Product;
Profile     := Param.Profile;
Position    := Param.Workoffset;
Direction   := Param.B;             /* 0 = inside, 1 = outside */
ProfileType := Param.C;             /* 0 = vent, 1 = frame, 2 = T-mullion */

h:=0;
p_height:=0;
if Product="0011441" then
{
 h:=0;
 p_height:=61;
}
else if Product="0011442" then
{
 h:=70;
 p_height:=70;
}
else if Product="1030211" then
{
 h:=0;
 p_height:=70;
}
else if Product="1030222" then
{
 h:=79;
 p_height:=79;
} 

/*7.5mm*/
Tool.Name := 'STD_SLOT';

Tool.X:=Position;
if Product="0011441" || Product="1030211" then
{
 Tool.Y:=24.5;
 Tool.Z:=2.3;
 Tool.P1X:=0;
 Tool.P1Y:=0;
 Tool.P1Z:=h;
 Tool.P2X:=0;
 Tool.P2Y:=63;
 Tool.P2Z:=h;
 Tool.P3X:=100;
 Tool.P3Y:=63;
 Tool.P3Z:=h;
}
else if Product="0011442" || Product="1030222" then
{
 Tool.Y:=24.5;
 Tool.Z:=2.3;
 Tool.P1X:=0;
 Tool.P1Y:=63;
 Tool.P1Z:=h;
 Tool.P2X:=100;
 Tool.P2Y:=63;
 Tool.P2Z:=h;
 Tool.P3X:=0;
 Tool.P3Y:=0;
 Tool.P3Z:=h;
}
Tool.Param['PAR1']:=0;
Tool.Param['PAR2']:=34;
Tool.Param['PAR3']:=6;
Tool.Tool:=2;
Tool.Apply();
