profile:=Action.Reference;
len:=Action.SelfLength;
p1:=Action.ReferencePos1;
p2:=Action.ReferencePos2;
d1:=4.5;
d2:=8.5;
edge_ofs_s:=59.5;
edge_ofs_g:=26;
edge_ofs_p:=36;
pos_s:=abs(p1-edge_ofs_s);
pos_g:=abs(p1-edge_ofs_g);
pos_p:=abs(p1-edge_ofs_p);
/*msgbox(p1,",",edge_ofs_g,",",pos_g);*/
/*msgbox(profile.profile.system);
msgbox(profile.profile.code);*/
if profile.profile.system="ES70A" then
{
  Machine.Do(profile,"MITRE_SCREW_ES70A",POS_OFFSET,pos_s,0,0);
  Machine.Do(profile,"MITRE_GLUE_ES70A",POS_OFFSET,pos_g,0,0);
  Machine.Do(profile,"MITRE_PIN_ES70A",POS_OFFSET,pos_p,0,0);
}
else if profile.profile.system="ES61" then
{
  Machine.Do(profile,"MITRE_SCREW_ES61",POS_OFFSET,pos_s,0,0);
  Machine.Do(profile,"MITRE_GLUE_ES61",POS_OFFSET,pos_g,0,0);
  Machine.Do(profile,"MITRE_PIN_ES61",POS_OFFSET,pos_p,0,0);
}
else
 msgbox("profile.profile.code is neither ES70A nor ES61");