ProfielParametersOpen := 0;

Reeks1 := GetParam("REEKS1");
Code1  := GetParam("PROFIELCODE1");

Reeks2 := GetParam("REEKS2");
Code2  := GetParam("PROFIELCODE2");

ProfielCombinaties                         := Combinations.Create();
ProfielParameters                          := Profiles.Create();

/* Combinatie kader en vleugel */
ProfielCombinaties.Code.Profile1.System    := Reeks1;
ProfielCombinaties.Code.Profile1.Code      := Code1;
ProfielCombinaties.Code.Profile2.System    := Reeks2;
ProfielCombinaties.Code.Profile2.Code      := Code2;
if ProfielCombinaties.Find() then Goto Gevonden;

/* Combinatie vleugel en kader */
ProfielCombinaties.Code.Profile1.System    := Reeks2;
ProfielCombinaties.Code.Profile1.Code      := Code2;
ProfielCombinaties.Code.Profile2.System    := Reeks1;
ProfielCombinaties.Code.Profile2.Code      := Code1;

if ProfielCombinaties.Find() then Goto Gevonden;

/* Zoeken met alternatief voor kaderprofiel uit profiel combinaties */

ProfielParameters.Code.System  			   := Reeks1;
ProfielParameters.Code.Profile			   := Code1;
ProfielParameters.Find();
If ProfielParameters.Combine.System != "" then
{
  ProfielCombinaties.Code.Profile1.System    := ProfielParameters.Combine.System;
  ProfielCombinaties.Code.Profile1.Code      := ProfielParameters.Combine.Profile;
  ProfielCombinaties.Code.Profile2.System    := Reeks2;
  ProfielCombinaties.Code.Profile2.Code      := Code2;
  if ProfielCombinaties.Find() then Goto Gevonden;
  ProfielCombinaties.Code.Profile1.System    := Reeks2;
  ProfielCombinaties.Code.Profile1.Code      := Code2;
  ProfielCombinaties.Code.Profile2.System    := ProfielParameters.Combine.System;
  ProfielCombinaties.Code.Profile2.Code      := ProfielParameters.Combine.Profile;
  if ProfielCombinaties.Find() then Goto Gevonden;
} 
  
/* Zoeken met alternatief voor vleugelprofiel uit profiel combinaties */

ProfielParameters.Code.System  			   := Reeks2;
ProfielParameters.Code.Profile			   := Code2;
ProfielParameters.Find();
If ProfielParameters.Combine.System != "" then
{
  ProfielCombinaties.Code.Profile1.System    := Reeks1;
  ProfielCombinaties.Code.Profile1.Code      := Code1;
  ProfielCombinaties.Code.Profile2.System    := ProfielParameters.Combine.System;
  ProfielCombinaties.Code.Profile2.Code      := ProfielParameters.Combine.Profile;
  if ProfielCombinaties.Find() then Goto Gevonden;
  ProfielCombinaties.Code.Profile1.System    := ProfielParameters.Combine.System;
  ProfielCombinaties.Code.Profile1.Code      := ProfielParameters.Combine.Profile;
  ProfielCombinaties.Code.Profile2.System    := Reeks1;
  ProfielCombinaties.Code.Profile2.Code      := Code1;
  if ProfielCombinaties.Find() then Goto Gevonden;
} 

/* Zoeken met alternatief voor beide profielen uit profiel combinaties */

ProfielParameters.Code.System  			     := Reeks1;
ProfielParameters.Code.Profile			     := Code1;
ProfielParametersOpen := 1;
ProfielParameters.Find();
If ProfielParameters.Combine.System != "" then
{
  Reeks3							         := ProfielParameters.Combine.System;
  Code3										 := ProfielParameters.Combine.Profile;
}
else
{
  msgbox("Profielcombinatie niet gevonden in bibliotheek voor profielen " + Reeks1 + " " + Code1 + " en " + Reeks2 + " " + Code2);
  goto Einde;
}
ProfielParameters.Code.System  			     := Reeks2;
ProfielParameters.Code.Profile			     := Code2;
ProfielParameters.Find();
If ProfielParameters.Combine.System != "" then
{
  Reeks4							         := ProfielParameters.Combine.System;
  Code4										 := ProfielParameters.Combine.Profile;
}
else
{
  msgbox("Profielcombinatie niet gevonden in bibliotheek voor profielen " + Reeks1 + " " + Code1 + " en " + Reeks2 + " " + Code2);
  goto Einde;
}
ProfielCombinaties.Code.Profile1.System    := Reeks3;
ProfielCombinaties.Code.Profile1.Code      := Code3;
ProfielCombinaties.Code.Profile2.System    := Reeks4;
ProfielCombinaties.Code.Profile2.Code      := Code4;
if ProfielCombinaties.Find() then Goto Gevonden;
ProfielCombinaties.Code.Profile1.System    := Reeks4;
ProfielCombinaties.Code.Profile1.Code      := Code4;
ProfielCombinaties.Code.Profile2.System    := Reeks3;
ProfielCombinaties.Code.Profile2.Code      := Code3;
if ProfielCombinaties.Find() then Goto Gevonden;
 
msgbox("Profielcombinatie niet gevonden in bibliotheek voor profielen " + Reeks1 + " " + Code1 + " en " + Reeks2 + " " + Code2);
  goto Einde;

@Gevonden:
Overlapping1                      		   := ProfielCombinaties.Overlap[0];

SetParam("OVERLAP",NumToStr(Overlapping1));

@Einde:
ProfielParameters.Free();
ProfielCombinaties.free();
