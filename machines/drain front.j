/* DRAIN FRONT */

Product     := Param.Product;
Profile     := Param.Profile;
Position    := Param.Workoffset;
Direction   := Param.B;             /* 0 = inside, 1 = outside */
ProfileType := Param.C;             /* 0 = vent, 1 = frame, 2 = T-mullion */

/*7.5mm*/
Tool.Name := 'STD_SLOT';

Tool.X:=Position;
Tool.Y:=24.5;
Tool.Z:=2.3;
Tool.P1X:=0;
Tool.P1Y:=63;
Tool.P1Z:=0;
Tool.P2X:=0;
Tool.P2Y:=0;
Tool.P2Z:=0;
Tool.P3X:=100;
Tool.P3Y:=0;
Tool.P3Z:=0;
Tool.Param['PAR1']:=0;
Tool.Param['PAR2']:=34;
Tool.Param['PAR3']:=6;
Tool.Tool:=2;
Tool.Apply();
