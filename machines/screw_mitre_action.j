profile:=Action.Reference;
len:=Action.SelfLength;
p1:=Action.ReferencePos1;
p2:=Action.ReferencePos2;
edge_ofs_s:=59.5;
edge_ofs_g:=26;
pos_s:=abs(p1-edge_ofs_s);
pos_g:=abs(p1-edge_ofs_g);
/*msgbox(profile.profile.system);
msgbox(profile.profile.code);*/
if profile.profile.system="ES70A" then
{
  Machine.Do(profile,"MITRE_SCREW_ES70A",Pos_OFFSET,pos_s,0,0);
  Machine.Do(profile,"MITRE_GLUE_ES70A",Pos_OFFSET,pos_g,0,0);
}
else if profile.profile.system="ES61" then
{
  Machine.Do(profile,"MITRE_SCREW_ES61",Pos_OFFSET,pos_s,0,0);
  Machine.Do(profile,"MITRE_GLUE_ES61",Pos_OFFSET,pos_g,0,0);
}
else
 msgbox("profile.profile.code is neither ES70A nor ES61");
