/* DRAIN FRONT */
Product     := Param.Product;
Profile     := Param.Profile;
Position    := Param.Workoffset;
Direction   := Param.B;             /* 0 = inside, 1 = outside */
ProfileType := Param.C;             /* 0 = vent, 1 = frame, 2 = T-mullion */

Tool.Name := 'DRAIN FRONT';
Tool.ProfiLav := True;
Tool.X := Position;
Tool.Apply();
