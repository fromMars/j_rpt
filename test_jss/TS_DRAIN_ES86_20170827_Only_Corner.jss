/**------------------------------------------------------------**/
/** FOR DRAIN_ES86	                               2015.02.28  **/
/** BY PAN QINGXI                                              **/
/**------------------------------------------------------------**/

open   := Action.Self;
frame  := open.parent;
target := Action.Reference;
ofs    := Action.ReferenceOfs1;
len    := Action.SelfLength;
len2   := target.length;
mop := 'V_DRAIN_SLIDING';
slide_cnt := 0;                /* number of vents in this opening */ 
triple_rail:=false;


counter := 0;
child_cnt := frame.childcount;
while counter < child_cnt do
{/*1*/
   child := frame.children[counter];
   if child.isframeopening then
      slide_cnt := slide_cnt +1;
   counter := counter +1
}/*1*/
/** EXclude SYSTEM**/
/*
if target.profile.system='ES61' then goto stop; 
if target.profile.system='ES52' then goto stop; 
if target.profile.system='ES40' then goto stop; 
if target.profile.system='ES70' then goto stop; 
if target.profile.system='ES70A' then goto stop; 
if target.profile.system='ES78' then goto stop; 
if target.profile.system='EF60' then goto stop; 
if target.profile.system='EC52' then goto stop; 
if target.profile.system='SC' then goto stop; 
if target.profile.system='ES144' then goto stop; 
if target.profile.system='ES152' then goto stop; 
if target.profile.system='ES116' then goto stop;
if target.profile.system='ESL' then goto stop;
if target.profile.system='HUAL' then goto stop;
if target.profile.system='ES100' then goto stop;
if target.profile.system='ES100A' then goto stop;
if target.profile.system='ES101' then goto stop;
if target.profile.system='EOSS' then goto stop;*/
if target.profile.system<>'ES86' then goto stop;

if open.isframeopening && open.childcount != 0 then

if !target.IsHorizontal then goto stop; 


if target.isventelement  then 
{/*1*/
   if (target.code.kind = ELMTKIND_VENTFRAME) then 
   mop := 'V_DRAIN_SLIDING';
   Drain_opt_corner              :=80;
   Drain_max_distance_one_hole   :=200;
   Drain_max_distance_no_hole    :=100;
   Drain_max_dist                :=1000;
   Drain_min_dist                :=250;


   If (Len < Drain_max_distance_no_hole ) Then  /* too, small no draining holes */
      goto Stop;

   If len < 250 then                            /* smaller then 550 two holes at 40 (straight down ) */
   {/*2*/
      Drain_opt_corner              :=40;
   }/*2*/

	  /* normal drain down rule*/
   If (Len < Drain_max_distance_one_hole) Then  /* One drain in the middle */
   {/*2*/
      Pos:=Len / 2;
      Machine.Do(target,mop,POS_OFFSET,Pos-ofs); /**ÉÈÅÅË®¿×**/
      goto Stop; 
   }/*2*/

   Pos:= Drain_opt_corner;
   Machine.Do(target,mop,POS_OFFSET,Pos-ofs); /**ÉÈÅÅË®¿×**/

   Rest_length:= Len -2 * Drain_opt_corner;
   extra_number:= Rest_length // Drain_max_dist;

   counter:=0;
   While (counter < extra_number) do
   {/*2*/
      Pos:=Drain_opt_corner+ (counter+1) * (Rest_length / (Extra_number + 1));
      Machine.Do(target,mop,POS_OFFSET,Pos-ofs); /**ÉÈÅÅË®¿×**/
      counter:=counter+1;
   }/*2*/
   Pos:= Len-Drain_opt_corner;
   Machine.Do(target,mop,POS_OFFSET,Pos-ofs); /**ÉÈÅÅË®¿×**/
}/*1*/

else
if target.IsOuterFrame && open.IsFrameOpening then
{/*1*/
   vent := open.children[0];
   if vent.isventpart then
   {/*2*/
      prof := profiles.create ();
      prof.code.system  := Target.profile.system;
      prof.code.profile := Target.profile.code;
      if !prof.find () then goto STOP;           /* profile not found in profile parameters */
      
      if prof.geometry [2][0] != 0 then          /* geometry field Y0=1 for monorail */
      triple_rail:=true;
      prof.free();

      if vent.kind = ventkind_slide || vent.kind = ventkind_fixed ||vent.kind = ventkind_liftslide   then  
      {/*3*/
       mop := 'B_DRAIN_SLIDING';   
 	     mop2:= 'A_DRAIN_SLIDING';
 	     mop3:= 'DRAIN_FRONT';
  
         if vent.link = 1 || (vent.link = 2 && triple_rail=true) then /* inner rail devider settings */
         {/*4*/
            Drain_opt_corner              :=100;
            Drain_max_dist                :=300;
            Drain_min_dist                :=250;
         }/*4*/
         if (vent.link = 2 && triple_rail=false) || (vent.link = 3 && triple_rail=true) then /* outer rail devider settings */
         {/*4*/
            Drain_opt_corner              :=100;
            Drain_max_dist                :=800;
            Drain_min_dist                :=250;
         }/*4*/
         
         pos1:=120;
         pos2:=len2-120;

         if vent.definition.contour[0].bar.Xb<100 then
         {
            Machine.Do(target,mop2,POS_OFFSET,pos1);
            Machine.Do(target,mop,POS_OFFSET,pos2);
         }       
    }/*3*/
   }/*2*/
}/*1*/
@Stop:
