unit UtilTableauBord;

interface

uses
  Classes,
  Hctrls,
  {$IFNDEF EAGLCLIENT}
  db,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  Fe_Main,
  {$ELSE}
  MaineAGL,
  {$ENDIF}
  sysutils,
  HEnt1,
  HMsgBox,
  UTOB,
  Entgc,
  AGLInit,
  Paramsoc,
  Ulog,
  AGLInitBTP;

type
	ToptPrevuAvanc = (OptDetailDev,OptDetailPBT,OptDetailETU,OptDetailBCE,OptGlobal,OptDetailFAC,OptDetailBu);
  TChoixPrevuAvanc = set of ToptPrevuAvanc;


procedure AjouteChampSupPrevu(TOBTmp : TOB; Nature : string);
procedure SetprevuAvance (TOBTMP : TOB; NaturePiece : string ; OptionChoixPrevuAvanc : TChoixPrevuAvanc; DateMvtFin : TdateTime; WherePiece : string='');
procedure AjouteChampSupRealiseDetail (TOBTmp : TOB);
procedure AjouteChampSupFactureDetail (TOBTmp : TOB);
procedure AjouteChampSupRestADep(TOBTMP : TOB);

Procedure MAJ_PremiereLigne(NomChampSup : String; TOBTMP : Tob);
Procedure Repartition_Eclatement(TypeMontant, NaturePres, TypeMt : String; Montant : Double; TOBTMP : Tob);
Procedure ChargeCumulPrevuFacture(TypeRes, NaturePiece : String; TOBTMP : TOB; MontantPa,MontantPr,MontantPV,TpsPrevu : Double);
Procedure Charge_Repartition_Eclatement(NaturePres : String; MontantPA, MontantPR, MontantPV : Double; TOBTMP : TOB);
Procedure ChargePrevu(TOBL, TOBTMP : TOB; Prefixe : String; Tps_Prevu : Double =0);

procedure TraiteOuvrageTBPlat (TOBTMP,TOBL,TOBOuvragePlat : TOB);

procedure AjouteChampsSupBu (TOBTMP : TOB);
procedure RepartitionConsommeBu (TOBE,TOBTMP : TOB);
procedure RepartitionPrevuAvanceBu (TOBTMP : TOB; CodeBu : string; PrevuPa,PrevuPr,PrevuPV,AvancePa,AvancePr,AvancePv : double);

//Modif FV1 :
procedure TraiteDetailOuvrageTBPlat (TOBTMP,TOBOUV : TOB);
procedure TraiteDetailOuvrageTB (TOBTMP,TOBOUV : TOB; Qte,QteDuDetail : double);

  //Modif FV : Dev. prioritaire DSL le 05/06/2012
var CoefFG_Param  : Double;
    TauxHoraire   : Double;
    //
implementation
uses UtilsTOB,UAvtVerdon;


procedure CumuleprevufactureAutre (TOBTMP: TOB; NaturePiece : string; MontantPa,MontantPr,MontantPV : double);
Var FactureAutre  : Double;
    TotalMoAutPA  : Double;
    TotalMoAutPR  : Double;
    TotalMoAutPV  : Double;
begin

  if (Pos(NaturePiece,'FPR;FBT;FBP;FBC;FAC;DAC')>0) or (Pos(NaturePiece,'ABT;ABC;AVC')>0) Then
  begin
     FactureAutre := TOBTMP.GetDouble('FACTURE_AUTRE');
     FactureAutre := FactureAutre + MontantPV;
     //
     TOBTMP.PutValue('FACTURE_AUTRE', FactureAutre);
  end
  else
  begin
     TotalMoAutPa   := TOBTMP.GetDouble('PREVU_'+NaturePiece+'_MOSALPA');
     TotalMoAutPa   := TotalMoAutPa + MontantPA;
     //
     TotalMoAutPR   := TOBTMP.GetDouble('PREVU_'+NaturePiece+'_MOSALPR');
     TotalMoAutPR   := TotalMoAutPR + MontantPR;
     //
     TotalMoAutPV   := TOBTMP.GetDouble('PREVU_'+NaturePiece+'_MOSALPV');
     TotalMoAutPV   := TotalMoAutPV + MontantPV;
     //
     TOBTMP.PutValue('PREVU_'+NaturePiece+'_AUTPA', TotalMoAutPa);
     TOBTMP.PutValue('PREVU_'+NaturePiece+'_AUTPR', TotalMoAutPR);
     TOBTMP.PutValue('PREVU_'+NaturePiece+'_AUTPV', TotalMoAutPV);
  end;

end;

procedure CumuleprevufactureSalarie (TOBTMP : TOB ; NaturePiece : string; MontantPa,MontantPr,MontantPV,TpsPrevu : double);
var FactureSalarie  : Double;
    TotalMoSalPA    : Double;
    TotalMoSalPR    : Double;
    TotalMoSalPV    : Double;
    TotalMoSal      : Double;
begin

  //FV1 : 21/11/2018 - FS#3376 - SCETEC - Les prévisionnels ne sont plus gérés dans les cumuls prévus
  if (Pos(NaturePiece,'FPR;FBT;FBP;FBC;FAC;DAC')>0) or (Pos(NaturePiece,'ABT;ABC;AVC')>0) Then
  begin
     FactureSalarie := TOBTMP.GetDouble('FACTURE_SALARIE');
     FactureSalarie := FactureSalarie + MontantPV;
     //
     TOBTMP.PutValue('FACTURE_SALARIE', FactureSalarie);
  end
  else
  begin
     TotalMoSalPa   := TOBTMP.GetDouble('PREVU_'+NaturePiece+'_MOSALPA');
     TotalMoSalPa   := TotalMoSalPa + MontantPA;
     //
     TotalMoSalPR   := TOBTMP.GetDouble('PREVU_'+NaturePiece+'_MOSALPR');
     TotalMoSalPR   := TotalMoSalPR + MontantPR;
     //
     TotalMoSalPV   := TOBTMP.GetDouble('PREVU_'+NaturePiece+'_MOSALPV');
     TotalMoSalPV   := TotalMoSalPV + MontantPV;
     //
     TotalMoSal     := TOBTMP.Getdouble('TPS_PREVU_'+NaturePiece+'_MOSAL');
     TotalMoSal     := TotalMoSal + TpsPRevu;
     //
     //FV1 - 16/11/2018 : FS#3368 - DELABOUDINIERE - Erreur sur Tableau de Bord sur contrat : Impossible de convertir un variant...
     TOBTMP.PutValue('PREVU_'+NaturePiece+'_MOSALPA', TotalMoSalPa);
     TOBTMP.PutValue('PREVU_'+NaturePiece+'_MOSALPR', TotalMoSalPR);
     TOBTMP.PutValue('PREVU_'+NaturePiece+'_MOSALPV', TotalMoSalPV);
     TOBTMP.PutValue('TPS_PREVU_'+NaturePiece+'_MOSAL', TotalMoSal);
     //FV1 - 16/11/2018 : FS#3368 - DELABOUDINIERE - Erreur sur Tableau de Bord sur contrat : Impossible de convertir un variant...
  end;

end;

Procedure Repartition_Eclatement(TypeMontant, NaturePres, TypeMt : String; Montant : Double; TOBTMP : Tob);
var MtAvant, MtApres : Double;
Begin

  if Not TobTMP.FieldExists(TypeMontant + NaturePres + TypeMt) then
  begin
    TOBTMP.AddChampSupValeur(TypeMontant + NaturePres + TypeMt, Montant);
    MAJ_PremiereLigne(TypeMontant + NaturePres + TypeMt, TOBTMP);
  end else
  begin
    MtAvant := TOBTMP.GetDouble(TypeMontant + NaturePres + TypeMt);
    MtAvant := Arrondi(MtAvant, V_PGI.OkDecV);
    MtApres := MtAvant + Montant;
    MtApres := Arrondi(MtApres, V_PGI.OkDecV);
    TOBTMP.PutValue(TypeMontant + NaturePres  + TypeMt,  MtApres);
  end;
end;

Procedure MAJ_PremiereLigne(NomChampSup : String; TOBTMP : TOB);
Begin
  // ATTENTION, on crée le champ sup dans la première ligne de la TOB principale pour que le champ soit présent dans les champs disponibles du tobviewer
  if TOBTMP.Parent.Detail.count = 0 Then Exit;
  if Not TOBTMP.Parent.Detail[0].FieldExists(NomChampSup) then
  begin
     TOBTMP.Parent.Detail[0].AddChampSupValeur(NomChampSup, 0.0);
  end;
end;


procedure CumuleprevufactureInterimaire (TOBTMP : TOB ; NaturePiece : string; MontantPa,MontantPr,MontantPV,TpsPrevu : double);
var FactureInterim  : Double;
    TotalMoIntPA    : Double;
    TotalMoIntPR    : Double;
    TotalMoIntPV    : Double;
    TotalMoInt      : Double;
begin

  //if (Pos(NaturePiece,'FAC;FBT;B00;FBP;FBC;DAC')>0) or (Pos(NaturePiece,'AVC;ABT;ABP;ABC')>0) Then
  if (Pos(NaturePiece,'FPR;FBT;FBP;FBC;FAC;DAC')>0) or (Pos(NaturePiece,'ABT;ABC;AVC')>0) Then
  begin
     FactureInterim := TOBTMP.GetDouble('FACTURE_INTERIM');
     FactureInterim := FactureInterim + MontantPV;
     //
     TOBTMP.PutValue('FACTURE_INTERIM', FactureInterim);
  end
  else
  begin
    //FV1 - 16/11/2018 : FS#3368 - DELABOUDINIERE - Erreur sur Tableau de Bord sur contrat : Impossible de convertir un variant...
    //
    TotalMoIntPa   := TOBTMP.GetDouble('PREVU_'+NaturePiece+'_MOINTPA');
    TotalMoIntPa   := TotalMoIntPa + MontantPA;
    //
    TotalMoIntPR   := TOBTMP.GetDouble('PREVU_'+NaturePiece+'_MOINTPR');
    TotalMoIntPR   := TotalMoIntPR + MontantPR;
    //
    TotalMoIntPV   := TOBTMP.GetDouble('PREVU_'+NaturePiece+'_MOINTPV');
    TotalMoIntPV   := TotalMoIntPV + MontantPV;
    //
    TotalMoInt     := TOBTMP.Getdouble('TPS_PREVU_'+NaturePiece+'_MOINT');
    TotalMoInt     := TotalMoInt + TpsPRevu;
    //
    TOBTMP.PutValue('PREVU_'+NaturePiece+'_MOINTPA', TotalMoIntPa);
    TOBTMP.PutValue('PREVU_'+NaturePiece+'_MOINTPR', TotalMoIntPR);
    TOBTMP.PutValue('PREVU_'+NaturePiece+'_MOINTPV', TotalMoIntPV);
    TOBTMP.PutValue('TPS_PREVU_'+NaturePiece+'_MOINT', TotalMoInt);
    //FV1 - 16/11/2018 : FS#3368 - DELABOUDINIERE - Erreur sur Tableau de Bord sur contrat : Impossible de convertir un variant...
  end;
end;

procedure CumuleprevufactureLocation (TOBTMP : TOB ; NaturePiece : string; MontantPa,MontantPr,MontantPV : double);
var FactureLocation: Double;
    TotalLocPA    : Double;
    TotalLocPR    : Double;
    TotalLocPV    : Double;
begin

  //if (Pos(NaturePiece ,'FAC;FBT;B00;FBP;FBC;DAC')>0) or (Pos(NaturePiece,'AVC;ABT;ABP;ABC')>0) Then
  if (Pos(NaturePiece,'FPR;FBT;FBP;FBC;FAC;DAC')>0) or (Pos(NaturePiece,'ABT;ABC;AVC')>0) Then
  begin
     FactureLocation := TOBTMP.GetDouble('FACTURE_LOCATION');
     FactureLocation := FactureLocation + MontantPV;
     //
     TOBTMP.PutValue('FACTURE_LOCATION', FactureLocation);
  end
  else
  begin
    //FV1 - 16/11/2018 : FS#3368 - DELABOUDINIERE - Erreur sur Tableau de Bord sur contrat : Impossible de convertir un variant...
    //
    TotalLocPa   := TOBTMP.GetDouble('PREVU_'+NaturePiece+'_LOCPA');
    TotalLocPa   := TotalLocPa + MontantPA;
    //
    TotalLocPR   := TOBTMP.GetDouble('PREVU_'+NaturePiece+'_LOCPR');
    TotalLocPR   := TotalLocPR + MontantPR;
    //
    TotalLocPV   := TOBTMP.GetDouble('PREVU_'+NaturePiece+'_LOCPV');
    TotalLocPV   := TotalLocPV + MontantPV;
    //
    TOBTMP.PutValue('PREVU_'+NaturePiece+'_LOCPA', TotalLocPA);
    TOBTMP.PutValue('PREVU_'+NaturePiece+'_LOCPR', TotalLocPR);
    TOBTMP.PutValue('PREVU_'+NaturePiece+'_LOCPV', TotalLocPV);
    //FV1 - 16/11/2018 : FS#3368 - DELABOUDINIERE - Erreur sur Tableau de Bord sur contrat : Impossible de convertir un variant...
  end;
end;

procedure CumuleprevufactureMateriel (TOBTMP : TOB; NaturePiece : string; MontantPa,MontantPr,MontantPV : double);
var FactureMateriel: Double;
    TotalMatPA    : Double;
    TotalMatPR    : Double;
    TotalMatPV    : Double;
begin

  //if (Pos(NaturePiece,'FAC,FBT;B00;FBP;FBC;DAC')>0) or (Pos(NaturePiece,'AVC,ABT;ABP;ABC')>0) Then
  if (Pos(NaturePiece,'FPR;FBT;FBP;FBC;FAC;DAC')>0) or (Pos(NaturePiece,'ABT;ABC;AVC')>0) Then
  begin
     FactureMateriel := TOBTMP.GetDouble('FACTURE_MATERIEL');
     FactureMateriel := FactureMateriel + MontantPV;
     //
     TOBTMP.PutValue('FACTURE_MATERIEL', FactureMateriel);
  end else
  begin
    //FV1 - 16/11/2018 : FS#3368 - DELABOUDINIERE - Erreur sur Tableau de Bord sur contrat : Impossible de convertir un variant...
    //
    TotalMatPa   := TOBTMP.GetDouble('PREVU_'+NaturePiece+'_MATPA');
    TotalMatPa   := TotalMatPa + MontantPA;
    //
    TotalMatPR   := TOBTMP.GetDouble('PREVU_'+NaturePiece+'_MATPR');
    TotalMatPR   := TotalMatPR + MontantPR;
    //
    TotalMatPV   := TOBTMP.GetDouble('PREVU_'+NaturePiece+'_MATPV');
    TotalMatPV   := TotalMatPV + MontantPV;
    //
    TOBTMP.PutValue('PREVU_'+NaturePiece+'_MATPA', TotalMatPa);
    TOBTMP.PutValue('PREVU_'+NaturePiece+'_MATPR', TotalMatPR);
    TOBTMP.PutValue('PREVU_'+NaturePiece+'_MATPV', TotalMatPV);
   //FV1 - 16/11/2018 : FS#3368 - DELABOUDINIERE - Erreur sur Tableau de Bord sur contrat : Impossible de convertir un variant...
  end;
end;

procedure CumuleprevufactureOutillage (TOBTMP : TOB; NaturePiece : string; MontantPa,MontantPr,MontantPV : double);
var FactureOutil  : Double;
    TotalOutPA    : Double;
    TotalOutPR    : Double;
    TotalOutPV    : Double;
begin

  //if (Pos(NaturePiece,'FAC;FBT;B00;FBP;FBC;DAC')>0) or (Pos(NaturePiece,'AVC;ABT;ABP;ABC')>0) Then
  if (Pos(NaturePiece,'FPR;FBT;FBP;FBC;FAC;DAC')>0) or (Pos(NaturePiece,'ABT;ABC;AVC')>0) Then
  begin
     FactureOutil := TOBTMP.GetDouble('FACTURE_OUTIL');
     FactureOutil := FactureOutil + MontantPV;
     //
     TOBTMP.PutValue('FACTURE_OUTIL', FactureOutil);
  end else
  begin
    //FV1 - 16/11/2018 : FS#3368 - DELABOUDINIERE - Erreur sur Tableau de Bord sur contrat : Impossible de convertir un variant...
    //
    TotalOutPa   := TOBTMP.GetDouble('PREVU_'+NaturePiece+'_OUTPA');
    TotalOutPa   := TotalOutPa + MontantPA;
    //
    TotalOutPR   := TOBTMP.GetDouble('PREVU_'+NaturePiece+'_OUTPR');
    TotalOutPR   := TotalOutPR + MontantPR;
    //
    TotalOutPV   := TOBTMP.GetDouble('PREVU_'+NaturePiece+'_OUTPV');
    TotalOutPV   := TotalOutPV + MontantPV;
    //
    TOBTMP.PutValue('PREVU_'+NaturePiece+'_OUTPA', TotalOutPa);
    TOBTMP.PutValue('PREVU_'+NaturePiece+'_OUTPR', TotalOutPR);
    TOBTMP.PutValue('PREVU_'+NaturePiece+'_OUTPV', TotalOutPV);
  end;
  //
end;

procedure CumuleprevufactureSousTraitance (TOBTMP : TOB; NaturePiece : string ;MontantPa,MontantPr,MontantPV,TpsPrevu : double);
var FactureSt : Double;
    TotalStPA : Double;
    TotalStPR : Double;
    TotalStPV : Double;
    TotalSt   : Double;
begin

  //if (Pos(NaturePiece,'FAC;FBT;B00;FBP;FBC;DAC')>0) or (Pos(NaturePiece,'AVC;ABT;ABP;ABC')>0) Then
  if (Pos(NaturePiece,'FPR;FBT;FBP;FBC;FAC;DAC')>0) or (Pos(NaturePiece,'ABT;ABC;AVC')>0) Then
  begin
     FactureSt := TOBTMP.GetDouble('FACTURE_ST');
     FactureSt := FactureSt + MontantPV;
     //
     TOBTMP.PutValue('FACTURE_ST', FactureSt);
  end
  else
  begin
    TotalStPa   := TOBTMP.GetDouble('PREVU_'+NaturePiece+'_STPA');
    TotalStPa   := TotalStPa + MontantPA;
    //
    TotalStPR   := TOBTMP.GetDouble('PREVU_'+NaturePiece+'_STPR');
    TotalStPR   := TotalStPR + MontantPR;
    //
    TotalStPV   := TOBTMP.GetDouble('PREVU_'+NaturePiece+'_STPV');
    TotalStPV   := TotalStPV + MontantPV;
    //
    TotalSt     := TOBTMP.Getdouble('TPS_PREVU_'+NaturePiece+'_ST');
    TotalSt     := TotalSt + TpsPRevu;
    //
    //FV1 - 23/10/2017 - FS#2704 - AVENEL - En tableau de bord message impossible de convertir type String en type Double
    TOBTMP.PutValue('PREVU_'+NaturePiece+'_STPA',TotalStPa);
    TOBTMP.PutValue('PREVU_'+NaturePiece+'_STPR',TotalStPR);
    TOBTMP.PutValue('PREVU_'+NaturePiece+'_STPV',TotalStPV);
    TOBTMP.PutValue('TPS_PREVU_'+NaturePiece+'_ST', TotalSt);
    //
  end;
end;

procedure CumuleprevufactureFourniture (TOBTMP : TOB; NaturePiece : string ; MontantPa,MontantPr,MontantPV : double);
Var FactureFourniture : Double;
    TotalFournPA : Double;
    TotalFournPR : Double;
    TotalFournPV : Double;
begin

  //if (Pos(NaturePiece,'FAC,FBT;B00;FBP;FBC;DAC')>0) or (Pos(NaturePiece,'AVC,ABT;ABP;ABC')>0) Then
  if (Pos(NaturePiece,'FPR;FBT;FBP;FBC;FAC;DAC')>0) or (Pos(NaturePiece,'ABT;ABC;AVC')>0) Then
  begin
     FactureFourniture := TOBTMP.GetDouble('FACTURE_FOURNITURE');
     FactureFourniture := FactureFourniture + MontantPV;
     //
     TOBTMP.PutValue('FACTURE_FOURNITURE', FactureFourniture);
  end
  else
  begin
    //FV1 - 23/10/2017 - FS#2704 - AVENEL - En tableau de bord message impossible de convertir type String en type Double
    TotalFournPa  := TOBTMP.GetDouble('PREVU_'+NaturePiece+'_FOURNPA');
    TotalFournPa  := TotalFournPa + MontantPA;
    //
    TotalFournPR  := TOBTMP.GetDouble('PREVU_'+NaturePiece+'_FOURNPR');
    TotalFournPR  := TotalFournPR + MontantPR;
    //
    TotalFournPV  := TOBTMP.GetDouble('PREVU_'+NaturePiece+'_FOURNPV');
    TotalFournPV  := TotalFournPV + MontantPV;
    //
    TOBTMP.PutValue('PREVU_'+NaturePiece+'_FOURNPA', TotalFournPA);
    TOBTMP.PutValue('PREVU_'+NaturePiece+'_FOURNPR', TotalFournPR);
    TOBTMP.PutValue('PREVU_'+NaturePiece+'_FOURNPV', TotalFournPV);
    //
  end;
end;

//FV1 : 04/02/2014 - FS#863 - SCETEC : Distinguer les frais au niveau des champs de prévisionnel
procedure CumuleprevufactureFrais (TOBTMP : TOB; NaturePiece : string ; MontantPa,MontantPr,MontantPV : double);
Var FactureFrais : Double;
    TotalFraisPA : Double;
    TotalFraisPR : Double;
    TotalFraisPV : Double;
begin

  //if (Pos(NaturePiece,'FBT;B00;FBP;FBC;DAC;FAC')>0) or (Pos(NaturePiece,'ABT;ABP;ABC;AVC')>0) Then
  if (Pos(NaturePiece,'FPR;FBT;FBP;FBC;FAC;DAC')>0) or (Pos(NaturePiece,'ABT;ABC;AVC')>0) Then
  begin
     FactureFrais := TOBTMP.GetDouble('FACTURE_FRAIS');
     FactureFrais := FactureFrais + MontantPV;
     //
     TOBTMP.PutValue('FACTURE_FRAIS', FactureFrais);
  end
  else
  begin
    //FV1 - 23/10/2017 - FS#2704 - AVENEL - En tableau de bord message impossible de convertir type String en type Double
    TotalFraisPa  := TOBTMP.GetDouble('PREVU_'+NaturePiece+'_FRAISPA');
    TotalFraisPa  := TotalFraisPa + MontantPA;
    //
    TotalFraisPR  := TOBTMP.GetDouble('PREVU_'+NaturePiece+'_FRAISPR');
    TotalFraisPR  := TotalFraisPR + MontantPR;
    //
    TotalFraisPV  := TOBTMP.GetDouble('PREVU_'+NaturePiece+'_FRAISPV');
    TotalFraisPV  := TotalFraisPV + MontantPV;
    //
    TOBTMP.PutValue('PREVU_'+NaturePiece+'_FRAISPA',TotalFraisPa);
    TOBTMP.PutValue('PREVU_'+NaturePiece+'_FRAISPR',TotalFraisPR);
    TOBTMP.PutValue('PREVU_'+NaturePiece+'_FRAISPV',TotalFraisPV);
    //
  end;
end;

procedure RepartitionPrevuAvanceBu (TOBTMP : TOB; CodeBu : string; PrevuPa,PrevuPr,PrevuPV,AvancePa,AvancePr,AvancePv : double);
begin
  if CodeBu = '' then CodeBu := 'HORSBU';
  TOBTMP.SetDouble('PREVU_'+CodeBU+'_PA', TOBTMP.GetDouble('PREVU_'+CodeBU+'_PA')+PrevuPa);
  TOBTMP.SetDouble('PREVU_'+CodeBU+'_PR', TOBTMP.GetDouble('PREVU_'+CodeBU+'_PR')+PrevuPr);
  TOBTMP.SetDouble('PREVU_'+CodeBU+'_PV', TOBTMP.GetDouble('PREVU_'+CodeBU+'_PV')+PrevuPv);
  //
  TOBTMP.SetDouble('AVANCE_'+CodeBU+'_PA', TOBTMP.GetDouble('AVANCE_'+CodeBU+'_PA')+AvancePa);
  TOBTMP.SetDouble('AVANCE_'+CodeBU+'_PR', TOBTMP.GetDouble('AVANCE_'+CodeBU+'_PR')+AvancePr);
  TOBTMP.SetDouble('AVANCE_'+CodeBU+'_PV', TOBTMP.GetDouble('AVANCE_'+CodeBU+'_PV')+AvancePv);

end;


procedure RepartitionConsommeBu (TOBE,TOBTMP : TOB);
var CodeBu : string;
begin
  CodeBu := TOBE.getString('BCO_BU');
  if CodeBu = '' then CodeBu := 'HORSBU';
  TOBTMP.SetDouble('CONSOMME_'+CodeBU+'_PA', TOBTMP.GetDouble('CONSOMME_'+CodeBU+'_PA')+TOBE.GetDOuble('ACHAT'));
  TOBTMP.SetDouble('CONSOMME_'+CodeBU+'_PR', TOBTMP.GetDouble('CONSOMME_'+CodeBU+'_PR')+TOBE.GetDOuble('REVIENT'));
  TOBTMP.SetDouble('CONSOMME_'+CodeBU+'_PV', TOBTMP.GetDouble('CONSOMME_'+CodeBU+'_PV')+TOBE.GetDOuble('VENTE'));
end;

procedure AjouteChampsSupBu (TOBTMP : TOB);
var QQ: TQuery;
begin
  QQ := OpenSQL('SELECT BBU_CODEBU FROM BTBU',True,-1,'',true);
  TRY
    if not QQ.eof then
    begin
      //
      TOBTMP.addchampsup('PREVU_HORSBU_PA', false); TOBTMP.PutValue('PREVU_HORSBU_PA', 0.0);
      TOBTMP.addchampsup('PREVU_HORSBU_PR', false); TOBTMP.PutValue('PREVU_HORSBU_PR', 0.0);
      TOBTMP.addchampsup('PREVU_HORSBU_PV', false); TOBTMP.PutValue('PREVU_HORSBU_PV', 0.0);
      TOBTMP.addchampsup('AVANCE_HORSBU_PA', false); TOBTMP.PutValue('AVANCE_HORSBU_PA', 0.0);
      TOBTMP.addchampsup('AVANCE_HORSBU_PR', false); TOBTMP.PutValue('AVANCE_HORSBU_PR', 0.0);
      TOBTMP.addchampsup('AVANCE_HORSBU_PV', false); TOBTMP.PutValue('AVANCE_HORSBU_PV', 0.0);
      TOBTMP.addchampsup('CONSOMME_HORSBU_PA', false); TOBTMP.PutValue('CONSOMME_HORSBU_PA', 0.0);
      TOBTMP.addchampsup('CONSOMME_HORSBU_PR', false); TOBTMP.PutValue('CONSOMME_HORSBU_PR', 0.0);
      TOBTMP.addchampsup('CONSOMME_HORSBU_PV', false); TOBTMP.PutValue('CONSOMME_HORSBU_PV', 0.0);
      //
      QQ.First;
      while not QQ.Eof do
      begin
        TOBTMP.addchampsup('PREVU_'+QQ.fields[0].AsString+'_PA', false); TOBTMP.PutValue('PREVU_'+QQ.fields[0].AsString+'_PA', 0.0);
        TOBTMP.addchampsup('PREVU_'+QQ.fields[0].AsString+'_PR', false); TOBTMP.PutValue('PREVU_'+QQ.fields[0].AsString+'_PR', 0.0);
        TOBTMP.addchampsup('PREVU_'+QQ.fields[0].AsString+'_PV', false); TOBTMP.PutValue('PREVU_'+QQ.fields[0].AsString+'_PV', 0.0);
        TOBTMP.addchampsup('AVANCE_'+QQ.fields[0].AsString+'_PA', false); TOBTMP.PutValue('AVANCE_'+QQ.fields[0].AsString+'_PA', 0.0);
        TOBTMP.addchampsup('AVANCE_'+QQ.fields[0].AsString+'_PR', false); TOBTMP.PutValue('AVANCE_'+QQ.fields[0].AsString+'_PR', 0.0);
        TOBTMP.addchampsup('AVANCE_'+QQ.fields[0].AsString+'_PV', false); TOBTMP.PutValue('AVANCE_'+QQ.fields[0].AsString+'_PV', 0.0);
        TOBTMP.addchampsup('CONSOMME_'+QQ.fields[0].AsString+'_PA', false); TOBTMP.PutValue('CONSOMME_'+QQ.fields[0].AsString+'_PA', 0.0);
        TOBTMP.addchampsup('CONSOMME_'+QQ.fields[0].AsString+'_PR', false); TOBTMP.PutValue('CONSOMME_'+QQ.fields[0].AsString+'_PR', 0.0);
        TOBTMP.addchampsup('CONSOMME_'+QQ.fields[0].AsString+'_PV', false); TOBTMP.PutValue('CONSOMME_'+QQ.fields[0].AsString+'_PV', 0.0);
        QQ.Next;
      end;
    end;
  FINALLY
    Ferme(QQ);
  END;
end;

procedure AjouteChampSupPrevu(TOBTmp : TOB; Nature : string);
begin

  //FV1 : 03/06/2013 - global à l'ensemble des documents du chantier
  TOBTMP.addchampsup('PREVU_'+Nature+'_PA', false); TOBTMP.PutValue('PREVU_'+Nature+'_PA', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_PR', false); TOBTMP.PutValue('PREVU_'+Nature+'_PR', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_PV', false); TOBTMP.PutValue('PREVU_'+Nature+'_PV', 0.0);

  {Depuis prévision}

  (* Autre *)
  TOBTMP.addchampsup('PREVU_'+Nature+'_AUTPA', false); TOBTMP.PutValue('PREVU_'+Nature+'_AUTPA', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_AUTPR', false); TOBTMP.PutValue('PREVU_'+Nature+'_AUTPR', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_AUTPV', false); TOBTMP.PutValue('PREVU_'+Nature+'_AUTPV', 0.0);

  (* fourniture *)
  TOBTMP.addchampsup('PREVU_'+Nature+'_FOURNPA', false); TOBTMP.PutValue('PREVU_'+Nature+'_FOURNPA', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_FOURNPR', false); TOBTMP.PutValue('PREVU_'+Nature+'_FOURNPR', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_FOURNPV', false); TOBTMP.PutValue('PREVU_'+Nature+'_FOURNPV', 0.0);

  //FVA : 04/02/2014 - FS#863 - SCETEC : Distinguer les frais au niveau des champs de prévisionnel
  (* frais *)
  TOBTMP.addchampsup('PREVU_'+Nature+'_FRAISPA', false); TOBTMP.PutValue('PREVU_'+Nature+'_FRAISPA', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_FRAISPR', false); TOBTMP.PutValue('PREVU_'+Nature+'_FRAISPR', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_FRAISPV', false); TOBTMP.PutValue('PREVU_'+Nature+'_FRAISPV', 0.0);

  (* Main oeuvre interne *)
  TOBTMP.addchampsup('PREVU_'+Nature+'_MOSALPA', false); TOBTMP.PutValue('PREVU_'+Nature+'_MOSALPA', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_MOSALPR', false); TOBTMP.PutValue('PREVU_'+Nature+'_MOSALPR', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_MOSALPV', false); TOBTMP.PutValue('PREVU_'+Nature+'_MOSALPV', 0.0);
  TOBTMP.addchampsup('TPS_PREVU_'+Nature+'_MOSAL', false); TOBTMP.PutValue('TPS_PREVU_'+Nature+'_MOSAL', 0.0);

  (* Main oeuvre interimaire *)
  TOBTMP.addchampsup('PREVU_'+Nature+'_MOINTPA', false); TOBTMP.PutValue('PREVU_'+Nature+'_MOINTPA', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_MOINTPR', false); TOBTMP.PutValue('PREVU_'+Nature+'_MOINTPR', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_MOINTPV', false); TOBTMP.PutValue('PREVU_'+Nature+'_MOINTPV', 0.0);
  TOBTMP.addchampsup('TPS_PREVU_'+Nature+'_MOINT', false); TOBTMP.PutValue('TPS_PREVU_'+Nature+'_MOINT', 0.0);

  (* Location *)
  TOBTMP.addchampsup('PREVU_'+Nature+'_LOCPA', false); TOBTMP.PutValue('PREVU_'+Nature+'_LOCPA', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_LOCPR', false); TOBTMP.PutValue('PREVU_'+Nature+'_LOCPR', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_LOCPV', false); TOBTMP.PutValue('PREVU_'+Nature+'_LOCPV', 0.0);

  (* Materiel *)
  TOBTMP.addchampsup('PREVU_'+Nature+'_MATPA', false); TOBTMP.PutValue('PREVU_'+Nature+'_MATPA', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_MATPR', false); TOBTMP.PutValue('PREVU_'+Nature+'_MATPR', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_MATPV', false); TOBTMP.PutValue('PREVU_'+Nature+'_MATPV', 0.0);

  (* Outillage *)
  TOBTMP.addchampsup('PREVU_'+Nature+'_OUTPA', false); TOBTMP.PutValue('PREVU_'+Nature+'_OUTPA', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_OUTPR', false); TOBTMP.PutValue('PREVU_'+Nature+'_OUTPR', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_OUTPV', false); TOBTMP.PutValue('PREVU_'+Nature+'_OUTPV', 0.0);

  (* Sous Traitance *)
  TOBTMP.addchampsup('PREVU_'+Nature+'_STPA', false); TOBTMP.PutValue('PREVU_'+Nature+'_STPA', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_STPR', false); TOBTMP.PutValue('PREVU_'+Nature+'_STPR', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_STPV', false); TOBTMP.PutValue('PREVU_'+Nature+'_STPV', 0.0);
  TOBTMP.addchampsup('TPS_PREVU_'+Nature+'_ST', false); TOBTMP.PutValue('TPS_PREVU_'+Nature+'_ST', 0.0);

  // Ajout BRL 7/11/2012 : Montants des frais détaillés chantier, généraux et répartis
  TOBTMP.addchampsup('PREVU_'+Nature+'_MONTANTFG', false); TOBTMP.PutValue('PREVU_'+Nature+'_MONTANTFG', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_MONTANTFC', false); TOBTMP.PutValue('PREVU_'+Nature+'_MONTANTFC', 0.0);
  TOBTMP.addchampsup('PREVU_'+Nature+'_MONTANTFR', false); TOBTMP.PutValue('PREVU_'+Nature+'_MONTANTFR', 0.0);
end;

procedure AjouteChampSupRestADep(TOBTmp : TOB);
begin

  TOBTMP.addchampsup('RESTEADEP_FINAFF', false);     TOBTMP.PutValue('RESTEADEP_FINAFF', 0.0);
  TOBTMP.addchampsup('RESTEADEP_MTRESTE', false);    TOBTMP.PutValue('RESTEADEP_MTRESTE', 0.0);
  TOBTMP.addchampsup('RESTEADEP_QTRESTE', false);    TOBTMP.PutValue('RESTEADEP_QTRESTE', 0.0);
  TOBTMP.addchampsup('RESTEADEP_DATEARRETEE', false);    TOBTMP.PutValue('RESTEADEP_DATEARRETEE', iDate1900);
  //
  TOBTMP.addchampsup('RESTADEP_AUT_MTRESTE', false); TOBTMP.PutValue('RESTADEP_AUT_MTRESTE', 0.0);
  TOBTMP.addchampsup('RESTADEP_AUT_FINAFF',  false); TOBTMP.PutValue('RESTADEP_AUT_FINAFF',  0.0);
  // Fourniture deja definie
  TOBTMP.addchampsup('RESTADEP_SAL_MTRESTE', false); TOBTMP.PutValue('RESTADEP_SAL_MTRESTE', 0.0);
  TOBTMP.addchampsup('RESTADEP_SAL_QTRESTE', false); TOBTMP.PutValue('RESTADEP_SAL_QTRESTE', 0.0);
  TOBTMP.addchampsup('RESTADEP_SAL_FINAFF',  false); TOBTMP.PutValue('RESTADEP_SAL_FINAFF',  0.0);
  //
  TOBTMP.addchampsup('RESTADEP_LOC_MTRESTE', false); TOBTMP.PutValue('RESTADEP_LOC_MTRESTE', 0.0);
  TOBTMP.addchampsup('RESTADEP_LOC_FINAFF',  false); TOBTMP.PutValue('RESTADEP_LOC_FINAFF',  0.0);
  //
  TOBTMP.addchampsup('RESTADEP_OUT_MTRESTE', false); TOBTMP.PutValue('RESTADEP_OUT_MTRESTE', 0.0);
  TOBTMP.addchampsup('RESTADEP_OUT_FINAFF',  false); TOBTMP.PutValue('RESTADEP_OUT_FINAFF',  0.0);
  //
  TOBTMP.addchampsup('RESTADEP_ST_MTRESTE',  false); TOBTMP.PutValue('RESTADEP_ST_MTRESTE', 0.0);
  TOBTMP.addchampsup('RESTADEP_ST_FINAFF',   false); TOBTMP.PutValue('RESTADEP_ST_FINAFF',  0.0);
  //
  TOBTMP.addchampsup('RESTADEP_MAT_MTRESTE', false); TOBTMP.PutValue('RESTADEP_MAT_MTRESTE', 0.0);
  TOBTMP.addchampsup('RESTADEP_MAT_FINAFF',  false); TOBTMP.PutValue('RESTADEP_MAT_FINAFF',  0.0);
  //
  TOBTMP.addchampsup('RESTADEP_INT_MTRESTE', false); TOBTMP.PutValue('RESTADEP_INT_MTRESTE', 0.0);
  TOBTMP.addchampsup('RESTADEP_INT_QTRESTE', false); TOBTMP.PutValue('RESTADEP_INT_QTRESTE', 0.0);
  TOBTMP.addchampsup('RESTADEP_INT_FINAFF',  false); TOBTMP.PutValue('RESTADEP_INT_FINAFF',  0.0);
  //
  TOBTMP.addchampsup('RESTADEP_ACH_MTRESTE', false); TOBTMP.PutValue('RESTADEP_ACH_MTRESTE', 0.0);
  TOBTMP.addchampsup('RESTADEP_ACH_FINAFF',  false); TOBTMP.PutValue('RESTADEP_ACH_FINAFF',  0.0);
  TOBTMP.addchampsup('RESTADEP_STK_MTRESTE', false); TOBTMP.PutValue('RESTADEP_STK_MTRESTE', 0.0);
  TOBTMP.addchampsup('RESTADEP_STK_FINAFF',  false); TOBTMP.PutValue('RESTADEP_STK_FINAFF',  0.0);
  TOBTMP.addchampsup('RESTADEP_FOU_MTRESTE', false); TOBTMP.PutValue('RESTADEP_FOU_MTRESTE', 0.0);
  TOBTMP.addchampsup('RESTADEP_FOU_FINAFF',  false); TOBTMP.PutValue('RESTADEP_FOU_FINAFF',  0.0);
  //
  TOBTMP.addchampsup('RESTADEP_FAC_MTRESTE', false); TOBTMP.PutValue('RESTADEP_FAC_MTRESTE', 0.0);
  TOBTMP.addchampsup('RESTADEP_FAC_FINAFF',  false); TOBTMP.PutValue('RESTADEP_FAC_FINAFF',  0.0);

end;

procedure AjouteChampSupRealiseDetail (TOBTmp : TOB);
begin
  TOBTMP.addchampsup('REALISE_AUTRE_PA', false); TOBTMP.PutValue('REALISE_AUTRE_PA', 0.0);
  TOBTMP.addchampsup('REALISE_AUTRE_PR', false); TOBTMP.PutValue('REALISE_AUTRE_PR', 0.0);
  TOBTMP.addchampsup('REALISE_AUTRE_PV', false); TOBTMP.PutValue('REALISE_AUTRE_PV', 0.0);
  
  // Fourniture deja definie
  TOBTMP.addchampsup('REALISE_SAL_PA', false); TOBTMP.PutValue('REALISE_SAL_PA', 0.0);
  TOBTMP.addchampsup('REALISE_SAL_PR', false); TOBTMP.PutValue('REALISE_SAL_PR', 0.0);
  TOBTMP.addchampsup('REALISE_SAL_PV', false); TOBTMP.PutValue('REALISE_SAL_PV', 0.0);
  //
  TOBTMP.addchampsup('REALISE_LOC_PA', false); TOBTMP.PutValue('REALISE_LOC_PA', 0.0);
  TOBTMP.addchampsup('REALISE_LOC_PR', false); TOBTMP.PutValue('REALISE_LOC_PR', 0.0);
  TOBTMP.addchampsup('REALISE_LOC_PV', false); TOBTMP.PutValue('REALISE_LOC_PV', 0.0);
  //
  TOBTMP.addchampsup('REALISE_OUTIL_PA', false); TOBTMP.PutValue('REALISE_OUTIL_PA', 0.0);
  TOBTMP.addchampsup('REALISE_OUTIL_PR', false); TOBTMP.PutValue('REALISE_OUTIL_PR', 0.0);
  TOBTMP.addchampsup('REALISE_OUTIL_PV', false); TOBTMP.PutValue('REALISE_OUTIL_PV', 0.0);
  //
  TOBTMP.addchampsup('REALISE_ST_PA', false); TOBTMP.PutValue('REALISE_ST_PA', 0.0);
  TOBTMP.addchampsup('REALISE_ST_PR', false); TOBTMP.PutValue('REALISE_ST_PR', 0.0);
  TOBTMP.addchampsup('REALISE_ST_PV', false); TOBTMP.PutValue('REALISE_ST_PV', 0.0);
  //
  TOBTMP.addchampsup('REALISE_MAT_PA', false); TOBTMP.PutValue('REALISE_MAT_PA', 0.0);
  TOBTMP.addchampsup('REALISE_MAT_PR', false); TOBTMP.PutValue('REALISE_MAT_PR', 0.0);
  TOBTMP.addchampsup('REALISE_MAT_PV', false); TOBTMP.PutValue('REALISE_MAT_PV', 0.0);
  //
  TOBTMP.addchampsup('REALISE_INTERIM_PA', false); TOBTMP.PutValue('REALISE_INTERIM_PA', 0.0);
  TOBTMP.addchampsup('REALISE_INTERIM_PR', false); TOBTMP.PutValue('REALISE_INTERIM_PR', 0.0);
  TOBTMP.addchampsup('REALISE_INTERIM_PV', false); TOBTMP.PutValue('REALISE_INTERIM_PV', 0.0);
  //
end;

procedure AjouteChampSupFactureDetail (TOBTmp : TOB);
begin
  TOBTMP.addchampsup('FACTURE_FOURNITURE', false);TOBTMP.PutValue('FACTURE_FOURNITURE', 0.0);
  TOBTMP.addchampsup('FACTURE_AUTRE', false);     TOBTMP.PutValue('FACTURE_AUTRE', 0.0);
  TOBTMP.addchampsup('FACTURE_SALARIE', false);   TOBTMP.PutValue('FACTURE_SALARIE', 0.0);
  TOBTMP.addchampsup('FACTURE_LOCATION', false);  TOBTMP.PutValue('FACTURE_LOCATION', 0.0);
  TOBTMP.addchampsup('FACTURE_OUTIL', false);     TOBTMP.PutValue('FACTURE_OUTIL', 0.0);
  TOBTMP.addchampsup('FACTURE_ST', false);        TOBTMP.PutValue('FACTURE_ST', 0.0);
  TOBTMP.addchampsup('FACTURE_MATERIEL', false);  TOBTMP.PutValue('FACTURE_MATERIEL', 0.0);
  TOBTMP.addchampsup('FACTURE_INTERIM', false);   TOBTMP.PutValue('FACTURE_INTERIM', 0.0);

  //FVA : 04/02/2014 - FS#863 - SCETEC : Distinguer les frais au niveau des champs de prévisionnel
  TOBTMP.addchampsup('FACTURE_FRAIS', false);     TOBTMP.PutValue('FACTURE_FRAIS', 0.0);

end;


procedure SetPrevuAvancePBTVerdon (TOBTMP : TOB; OptionChoixPrevuAvanc : TChoixPrevuAvanc; DateMvtFin : TdateTime);
var req         : string;
		QQ          : TQuery;
    TypeRes     : String;
    NaturePres  : String;
    TypeArticle : String;
    NumAvc      : Integer;
    CodeBu : string;
    MontantPA, MontantPr,MontantPV, MontantFG, MontantFC, MontantFR, AvancePa,AvancePr,AvancePv,TpsPrevu,TpsAvance : double;
    AvancePACum,AvancePrCum,AvancePVCum : Double;
begin
  NumAvc := AVCVerdon.FindNumAvc ( TOBTMP.GetValue('BCO_AFFAIRE'),DateMvtFin);
  // Récupération du prévu
  Req := 'SELECT GL_AFFAIRE,GL_TYPEARTICLE,BNP_TYPERESSOURCE,BNP_NATUREPRES,GLC_BU,  '+
         'GL_QTEFACT*GL_DPA AS ACHAT, '+
  			 'GL_QTEFACT*GL_DPR AS REVIENT, '+
         'GL_TOTALHTDEV AS VENTE, ' +
         'GL_MONTANTFG AS MONTANTFG, ' +
         'GL_MONTANTFC AS MONTANTFC, ' +
         'GL_MONTANTFR AS MONTANTFR, ' +
         'GL_QTEFACT AS TPS_PREVU';
  // .. et de l'avance
  Req := Req +',GL_QTEFACT,GL_DPA,GL_DPR,BLF_POURCENTAVANC,BLF_MTSITUATION,BLF_QTECUMULEFACT,BLF_MTCUMULEFACT ';
  (* - -----
  Req := Req + ',(GL_QTEFACT*GL_DPA) * (BLF_POURCENTAVANC/100) AS AVANCEPA'+
  			   		 ',(GL_QTEFACT*GL_DPR) * (BLF_POURCENTAVANC/100) AS AVANCEPR'+
  			   		 ',BLF_MTSITUATION AS AVANCEPV'+
               ',BLF_QTECUMULEFACT AS TPS_AVANCE' ;
  ------ *)
  Req := Req + 'FROM LIGNE ' +
               'LEFT JOIN LIGNECOMPL ON GL_NATUREPIECEG=GLC_NATUREPIECEG AND GL_SOUCHE=GLC_SOUCHE AND GL_NUMERO=GLC_NUMERO AND GL_INDICEG=GLC_INDICEG AND GL_NUMORDRE=GLC_NUMORDRE '+  
               'LEFT JOIN LIGNEFAC ON GL_NATUREPIECEG=BLF_NATUREPIECEG AND GL_SOUCHE=BLF_SOUCHE AND GL_NUMERO=BLF_NUMERO AND GL_INDICEG=BLF_INDICEG AND GL_NUMORDRE=BLF_NUMORDRE AND BLF_NUMAVC='+IntToStr(NumAvc)+' '+
  						 'LEFT JOIN ARTICLE ON GA_ARTICLE=GL_ARTICLE '+
               'LEFT JOIN NATUREPREST ON BNP_NATUREPRES=GA_NATUREPRES '+
               'WHERE GL_NATUREPIECEG = "PBT" AND GL_TYPELIGNE LIKE "AR%" AND GL_AFFAIRE="' + TOBTMP.GetValue('BCO_AFFAIRE') + '"';
  QQ := OpenSql (req,True);

  while not QQ.eof do
  begin
    TypeRes   := QQ.findfield('BNP_TYPERESSOURCE').AsString;
    NaturePres:= QQ.findfield('BNP_NATUREPRES').AsString;
    //
    //FV1 : 04/02/2014 - FS#863 - SCETEC : Distinguer les frais au niveau des champs de prévisionnel
    TypeArticle := QQ.findfield('GL_TYPEARTICLE').AsString;
    If TypeArticle <> 'FRA' then
    begin
      //FV1 : Si pas de nature de prestation alors sur Nature 'Fournitures'
      if TypeRes = '' then TypeRes := 'FOU';
      if NaturePres = '' then NaturePres := 'FOURNITURES';
    end
    else
    begin
      //FV1 : pas de nature de prestation alors sur Nature 'Frais'
      if TypeRes = ''    then TypeRes := 'FRA';
      if NaturePres = '' then NaturePres := 'FRAIS';
    end;
    //
    CodeBu := QQ.findfield('GLC_BU').AsString;
    MontantPA := QQ.findfield('ACHAT').AsFloat;
    MontantPR := QQ.findfield('REVIENT').AsFloat;
    MontantPV := QQ.findfield('VENTE').AsFloat;
    MontantFG := QQ.findfield('MONTANTFG').AsFloat;
    MontantFC := QQ.findfield('MONTANTFC').AsFloat;
    MontantFR := QQ.findfield('MONTANTFR').AsFloat;
    //
    AvancePA  := (QQ.findfield('GL_QTEFACT').AsFloat*QQ.findfield('GL_DPA').AsFloat) * (QQ.findfield('BLF_POURCENTAVANC').AsFloat/100);
    AvancePR  := (QQ.findfield('GL_QTEFACT').AsFloat*QQ.findfield('GL_DPR').AsFloat) * (QQ.findfield('BLF_POURCENTAVANC').AsFloat/100);
    AvancePV  := QQ.findfield('BLF_MTCUMULEFACT').AsFloat;
    AvancePACum := arrondi(QQ.findfield('BLF_QTECUMULEFACT').AsFloat*QQ.findfield('GL_DPA').AsFloat,V_PGI.okdecV);
    AvancePRCum := arrondi(QQ.findfield('BLF_QTECUMULEFACT').AsFloat*QQ.findfield('GL_DPR').AsFloat,V_PGI.okdecV);
    AvancePVCum := QQ.findfield('BLF_MTCUMULEFACT').AsFloat;
    //
    TpsPrevu  := QQ.findfield('TPS_PREVU').AsFloat;
    TpsAvance := QQ.findfield('BLF_QTECUMULEFACT').AsFloat;

    TOBTMP.PutValue('PREVUPA', TOBTMP.GetValue('PREVUPA') + MontantPA);
    TOBTMP.PutValue('PREVUPR', TOBTMP.GetValue('PREVUPR') + MontantPR);
    // BRL 13/08 : ajout cumul montant Frais pour PBT
    TOBTMP.PutValue('PREVU_MONTANTFG', TOBTMP.GetValue('PREVU_MONTANTFG') + MontantFG);
    TOBTMP.PutValue('PREVU_MONTANTFC', TOBTMP.GetValue('PREVU_MONTANTFC') + MontantFC);
    TOBTMP.PutValue('PREVU_MONTANTFR', TOBTMP.GetValue('PREVU_MONTANTFR') + MontantFR);
    TOBTMP.PutValue('PREVUPV', TOBTMP.GetValue('PREVUPV') + MontantPV);

    TOBTMP.PutValue('AVANCEPA', TOBTMP.GetValue('AVANCEPA') + AvancePa);
    TOBTMP.PutValue('AVANCEPR', TOBTMP.GetValue('AVANCEPR') + AvancePr);
    TOBTMP.PutValue('AVANCEPV', TOBTMP.GetValue('AVANCEPV') + AvancePV);

    if TypeRes = 'SAL' then
    begin
      TOBTMP.PutValue('TPS_PREVU', TOBTMP.GetValue('TPS_PREVU')+ TpsPrevu);
      TOBTMP.PutValue('TPS_AVANCE', TOBTMP.GetValue('TPS_AVANCE') + TpsAvance);
    end;

    if OptDetailBu in OptionChoixPrevuAvanc then
    begin
      RepartitionPrevuAvanceBu (TOBTMP,CodeBu,Montantpa,MontantPr,MontantPV,AvancePaCum,AvancePrCum,AvancePvCum);
    end;
    

    if OptDetailPBT in OptionChoixPrevuAvanc then
    begin
      // BRL 13/08 : Mise à jour montants de frais par nature et prévus totaux pour PBT
      TOBTMP.PutValue('PREVU_PBT_MONTANTFG', TOBTMP.GetValue('PREVU_MONTANTFG'));
      TOBTMP.PutValue('PREVU_PBT_MONTANTFC', TOBTMP.GetValue('PREVU_MONTANTFC'));
      TOBTMP.PutValue('PREVU_PBT_MONTANTFR', TOBTMP.GetValue('PREVU_MONTANTFR'));
      TOBTMP.PutValue('PREVU_PBT_PA', TOBTMP.GetValue('PREVUPA'));
      TOBTMP.PutValue('PREVU_PBT_PR', TOBTMP.GetValue('PREVUPR'));
      TOBTMP.PutValue('PREVU_PBT_PV', TOBTMP.GetValue('PREVUPV'));
      //
      ChargeCumulPrevuFacture(TypeRes, 'PBT', TOBTMP,MontantPa,MontantPr,MontantPV,TpsPrevu);
    end;
    //
    if (TypeRes <> 'FOU') and (TOBTMP.GetValue('CEclatNatPrest') = 'X') then
    begin
      Charge_Repartition_Eclatement('PBT_'+ TypeRes + '_' + NaturePres,MontantPA, MontantPR, MontantPV, TOBTMP)
    end;


    QQ.next;
  end;
  Ferme(QQ);
end;

procedure SetPrevuAvancePBT (TOBTMP : TOB; OptionChoixPrevuAvanc : TChoixPrevuAvanc);
var req         : string;
		QQ          : TQuery;
    TypeRes     : String;
    NaturePres  : String;
    TypeArticle : String;
    MontantPA, MontantPr,MontantPV, MontantFG, MontantFC, MontantFR, AvancePa,AvancePr,AvancePv,TpsPrevu,TpsAvance : double;
begin

  // Récupération du prévu
  Req := 'SELECT BNP_TYPERESSOURCE,BNP_NATUREPRES, GL_TYPEARTICLE, '+
         'SUM(GL_QTEFACT*GL_DPA) AS ACHAT, '+
  			 'SUM(GL_QTEFACT*GL_DPR) AS REVIENT, '+
         'SUM(GL_TOTALHTDEV) AS VENTE, ' +
         'SUM(GL_MONTANTFG) AS MONTANTFG, ' +
         'SUM(GL_MONTANTFC) AS MONTANTFC, ' +
         'SUM(GL_MONTANTFR) AS MONTANTFR, ' +
         'SUM(GL_QTEFACT) AS TPS_PREVU';
  // .. et de l'avance
  Req := Req + ',SUM(GL_QTEPREVAVANC*GL_DPA) AS AVANCEPA'+
  			   		 ',SUM(GL_QTEPREVAVANC*GL_DPR) AS AVANCEPR'+
               ',SUM(GL_TOTALHTDEV*(GL_POURCENTAVANC/100.0)) AS AVANCEPV'+
               ',SUM(GL_QTEPREVAVANC) AS TPS_AVANCE' ;
  Req := Req + ' FROM LIGNE ' +
  						 'LEFT JOIN ARTICLE ON GA_ARTICLE=GL_ARTICLE '+
               'LEFT JOIN NATUREPREST ON BNP_NATUREPRES=GA_NATUREPRES '+
               'WHERE GL_NATUREPIECEG = "PBT" AND GL_TYPELIGNE LIKE "AR%" AND GL_AFFAIRE="' + TOBTMP.GetValue('BCO_AFFAIRE') + '"' +
               'GROUP BY GL_AFFAIRE,GL_TYPEARTICLE,BNP_TYPERESSOURCE,BNP_NATUREPRES';
  QQ := OpenSql (req,True);

  while not QQ.eof do
  begin
    TypeRes   := QQ.findfield('BNP_TYPERESSOURCE').AsString;
    NaturePres:= QQ.findfield('BNP_NATUREPRES').AsString;
    //
    //FV1 : 04/02/2014 - FS#863 - SCETEC : Distinguer les frais au niveau des champs de prévisionnel
    TypeArticle := QQ.findfield('GL_TYPEARTICLE').AsString;
    If TypeArticle <> 'FRA' then
    begin
      //FV1 : Si pas de nature de prestation alors sur Nature 'Fournitures'
      if TypeRes = '' then TypeRes := 'FOU';
      if NaturePres = '' then NaturePres := 'FOURNITURES';
    end
    else
    begin
      //FV1 : pas de nature de prestation alors sur Nature 'Frais'
      if TypeRes = ''    then TypeRes := 'FRA';
      if NaturePres = '' then NaturePres := 'FRAIS';
    end;
    //
    MontantPA := QQ.findfield('ACHAT').AsFloat;
    MontantPR := QQ.findfield('REVIENT').AsFloat;
    MontantPV := QQ.findfield('VENTE').AsFloat;
    MontantFG := QQ.findfield('MONTANTFG').AsFloat;
    MontantFC := QQ.findfield('MONTANTFC').AsFloat;
    MontantFR := QQ.findfield('MONTANTFR').AsFloat;
    //
    AvancePA  := QQ.findfield('AVANCEPA').AsFloat;
    AvancePR  := QQ.findfield('AVANCEPR').AsFloat;
    AvancePV  := QQ.findfield('AVANCEPV').AsFloat;
    //
    TpsPrevu  := QQ.findfield('TPS_PREVU').AsFloat;
    TpsAvance := QQ.findfield('TPS_AVANCE').AsFloat;

    TOBTMP.PutValue('PREVUPA', TOBTMP.GetValue('PREVUPA') + MontantPA);
    TOBTMP.PutValue('PREVUPR', TOBTMP.GetValue('PREVUPR') + MontantPR);
    // BRL 13/08 : ajout cumul montant Frais pour PBT
    TOBTMP.PutValue('PREVU_MONTANTFG', TOBTMP.GetValue('PREVU_MONTANTFG') + MontantFG);
    TOBTMP.PutValue('PREVU_MONTANTFC', TOBTMP.GetValue('PREVU_MONTANTFC') + MontantFC);
    TOBTMP.PutValue('PREVU_MONTANTFR', TOBTMP.GetValue('PREVU_MONTANTFR') + MontantFR);
    TOBTMP.PutValue('PREVUPV', TOBTMP.GetValue('PREVUPV') + MontantPV);

    TOBTMP.PutValue('AVANCEPA', TOBTMP.GetValue('AVANCEPA') + AvancePa);
    TOBTMP.PutValue('AVANCEPR', TOBTMP.GetValue('AVANCEPR') + AvancePr);
    TOBTMP.PutValue('AVANCEPV', TOBTMP.GetValue('AVANCEPV') + AvancePV);

    if TypeRes = 'SAL' then
    begin
      TOBTMP.PutValue('TPS_PREVU', TOBTMP.GetValue('TPS_PREVU')+ TpsPrevu);
      TOBTMP.PutValue('TPS_AVANCE', TOBTMP.GetValue('TPS_AVANCE') + TpsAvance);
    end;

    if OptDetailPBT in OptionChoixPrevuAvanc then
    begin
      // BRL 13/08 : Mise à jour montants de frais par nature et prévus totaux pour PBT
      TOBTMP.PutValue('PREVU_PBT_MONTANTFG', TOBTMP.GetValue('PREVU_MONTANTFG'));
      TOBTMP.PutValue('PREVU_PBT_MONTANTFC', TOBTMP.GetValue('PREVU_MONTANTFC'));
      TOBTMP.PutValue('PREVU_PBT_MONTANTFR', TOBTMP.GetValue('PREVU_MONTANTFR'));
      TOBTMP.PutValue('PREVU_PBT_PA', TOBTMP.GetValue('PREVUPA'));
      TOBTMP.PutValue('PREVU_PBT_PR', TOBTMP.GetValue('PREVUPR'));
      TOBTMP.PutValue('PREVU_PBT_PV', TOBTMP.GetValue('PREVUPV'));
      //
      ChargeCumulPrevuFacture(TypeRes, 'PBT', TOBTMP,MontantPa,MontantPr,MontantPV,TpsPrevu);
    end;
    //
    if (TypeRes <> 'FOU') and (TOBTMP.GetValue('CEclatNatPrest') = 'X') then
    begin
      Charge_Repartition_Eclatement('PBT_'+ TypeRes + '_' + NaturePres,MontantPA, MontantPR, MontantPV, TOBTMP)
    end;


    QQ.next;
  end;
  Ferme(QQ);
end;

//Mise à jour des prevu par type de ressource ou Nature de Prestation.
Procedure  Charge_Repartition_Eclatement(NaturePres : String; MontantPA, MontantPR, MontantPV : Double; TOBTMP : TOB);
Begin

    Repartition_Eclatement('PREVU_',  NaturePres, '_PA', MontantPA, TOBTMP);
    Repartition_Eclatement('PREVU_',  NaturePres, '_PR', MontantPR, TOBTMP);
    Repartition_Eclatement('PREVU_',  NaturePres, '_PV', MontantPV, TOBTMP);

end;

Procedure ChargeCumulPrevuFacture(TypeRes, NaturePiece : String; TOBTMP : TOB; MontantPa,MontantPr,MontantPV,TpsPrevu : Double);
begin

  if TypeRes = 'SAL' then           // prevu salarie
  begin
    //Modif FV : Dev. prioritaire DSL le 05/06/2012
    if (CoefFG_Param <> 0) then
    begin
      MontantPA := Arrondi(TauxHoraire * TpsPrevu,V_PGI.OkDecV);
      MontantPR := Arrondi(MontantPA * CoefFG_Param, V_PGI.OkDecV);
    end;
    CumuleprevufactureSalarie (TOBTMP,NaturePiece,MontantPa,MontantPr,MontantPV,TpsPrevu);
  end
  else if TypeRes = 'AUT' then      // prevu autre
    CumuleprevufactureAutre (TOBTMP,NaturePiece,MontantPa,MontantPr,MontantPV)
  else if TypeRes = 'INT' then     // prevu interimaire
    CumuleprevufactureInterimaire (TOBTMP,NaturePiece,MontantPa,MontantPr,MontantPV,TpsPrevu)
  else
  if TypeRes = 'LOC' then           // prevu location
    CumuleprevufactureLocation (TOBTMP,NaturePiece,MontantPa,MontantPr,MontantPV)
  else
  if TypeRes = 'MAT' then           // prevu materiel
    CumuleprevufactureMateriel (TOBTMP,NaturePiece,MontantPa,MontantPr,MontantPV)
  else if TypeRes = 'OUT' then      // prevu outillage
    CumuleprevufactureOutillage (TOBTMP,NaturePiece,MontantPa,MontantPr,MontantPV)
  else if TypeRes = 'ST' then       // prevu sous traitance
    CumuleprevufactureSousTraitance (TOBTMP,NaturePiece,MontantPa,MontantPr,MontantPV,TpsPrevu)
  else if TypeRes = 'FOU' then       // prevu fourniture
    CumuleprevufactureFourniture (TOBTMP,NaturePiece,MontantPa,MontantPr,MontantPV)
  else
    CumuleprevufactureFrais (TOBTMP,NaturePiece,MontantPa,MontantPr,MontantPV);
end;

procedure ChargelesLignesTB (TOBLIgne,TOBPiece : TOB);
var Req : string;
		QQ : TQuery;
begin
	Req := 'SELECT BNP_TYPERESSOURCE,BNP_NATUREPRES,GL_NATUREPIECEG,GL_SOUCHE, GL_NUMERO,'   +
         'GL_INDICEG,GL_NUMLIGNE,GL_TYPEARTICLE,GL_TYPENOMENC,GL_ARTICLE,'  +
         'GL_NUMORDRE,GL_INDICENOMEN,GL_QTEFACT,GL_DPA,GL_DPR,' +
         'GL_MONTANTFC, GL_MONTANTFG, GL_MONTANTFR, GL_TOTALHTDEV,GL_PUHTDEV, GLC_NATURETRAVAIL FROM LIGNE ' +
  			 'LEFT JOIN ARTICLE ON GA_ARTICLE=GL_ARTICLE ' +
         'LEFT JOIN NATUREPREST ON BNP_NATUREPRES=GA_NATUREPRES ' +
         'LEFT JOIN LIGNECOMPL ON GLC_NATUREPIECEG=GL_NATUREPIECEG AND ' +
         'GLC_SOUCHE=GL_SOUCHE AND GLC_NUMERO=GL_NUMERO AND ' +
         'GLC_INDICEG=GL_INDICEG AND GLC_NUMORDRE=GL_NUMORDRE ' +
         'WHERE '+
         'GL_TYPELIGNE="ART" ' + // on ne prend que les lignes de type article (les commentaires...pfff)
         'AND GL_NATUREPIECEG = "'+TOBPiece.getValue('GP_NATUREPIECEG')+ '" '+
         'AND GL_SOUCHE="' + TOBPiece.GetValue('GP_SOUCHE') + '" ' +
         'AND GL_NUMERO=' + IntToStr(TOBPiece.GetValue('GP_NUMERO')) + ' '+
         'AND GL_INDICEG=' + IntToStr(TOBPiece.GetValue('GP_INDICEG')) + ' '+
         'AND GL_TOTALHTDEV <> 0 ' +
         'ORDER BY GL_NUMLIGNE';
  QQ := OpenSql (req,true);
  TOBLigne.loadDetailDb ('LIGNE','','',QQ,false);
  ferme (QQ);
end;

procedure ConstitueOuvragesTB (Prefixe : String; TOBPiece,TOBOuvrage,TOBLocOuvrage : TOB);
var Lig : integer;
		indice : integer;
    TOBLig ,TOBNOuv, TOBPere, TOBnewDet,TOBL : TOB;
    LigneN1,LigneN2,LigneN3,LigneN4,LigneN5 : integer;
    Table : String;
begin
	// Initialisation
	Lig := 0;
  TOBNouv := nil;

  if Prefixe = 'BLO' then
    Table := 'LIGNEOUV'
  else
    Table := 'LIGNEOUVPLAT';
  //
	for Indice := 0 TO TOBLocOuvrage.detail.count -1 do
  begin
    TOBLig := TOBLocOuvrage.detail[Indice];
    if TOBLIg.getValue(Prefixe + '_NUMLIGNE') <> Lig then
    begin
    	// rupture sur N° de ligne --> donc nouvel ouvrage
      TOBNOuv := TOB.create ('NEW OUV',TOBOuvrage,-1);
      Lig := TOBLIg.getValue(Prefixe + '_NUMLIGNE');
      TOBL := TOBPiece.findFirst(['GL_NUMLIGNE'],[Lig],true);
      if TOBL<> nil then TOBL.putValue('GL_INDICENOMEN',TOBOuvrage.detail.count);
    end;
    LigneN1 := TOBLig.GetValue(Prefixe + '_N1');
    LigneN2 := TOBLig.GetValue(Prefixe + '_N2');
    LigneN3 := TOBLig.GetValue(Prefixe + '_N3');
    LigneN4 := TOBLig.GetValue(Prefixe + '_N4');
    LigneN5 := TOBLig.GetValue(Prefixe + '_N5');

    if LigneN5 > 0 then
    begin
			TOBPere:=TOBNOuv.FindFirst([Prefixe + '_NUMLIGNE',Prefixe + '_N1',Prefixe + '_N2',Prefixe + '_N3',Prefixe + '_N4',Prefixe + '_N5'],[Lig,LigneN1,LigneN2,LigneN3,LigneN4,0],True) ;
    end else
    if LigneN4 > 0 then
    begin
    	TOBPere:=TOBNOuv.FindFirst([Prefixe + '_NUMLIGNE',Prefixe + '_N1',Prefixe + '_N2',Prefixe + '_N3',Prefixe + '_N4',Prefixe + '_N5'],[Lig,LigneN1,LigneN2,LigneN3,0,0],True) ;
    end else
    if LigneN3 > 0 then
    begin
    	TOBPere:=TOBNOuv.FindFirst([Prefixe + '_NUMLIGNE',Prefixe + '_N1',Prefixe + '_N2',Prefixe + '_N3',Prefixe + '_N4',Prefixe + '_N5'],[Lig,LigneN1,LigneN2,0,0,0],True) ;
    end else
    if LigneN2 > 0 then
    begin
      TOBPere:=TOBNOuv.FindFirst([Prefixe + '_NUMLIGNE',Prefixe + '_N1',Prefixe + '_N2',Prefixe + '_N3',Prefixe + '_N4',Prefixe + '_N5'],[Lig,LigneN1,0,0,0,0],True) ;
    end else
    begin
    	TOBPere:=TOBNOuv;
    end;

    if TOBPere<>Nil then
    BEGIN
       TOBNewDet:=TOB.Create(Table,TOBPere,-1) ;
       TOBNewDet.Dupliquer(TOBLig,False,True) ;
    END;
  end;
end;

procedure ChargelesOuvragesTB (TOBOuvrage,TOBPiece : TOB);
var Req : string;
		QQ : TQuery;
    TOBLocOuvrage : TOB;
begin
	TOBLocOuvrage := TOB.Create ('OUV LU',nil,-1);
	Req := 'SELECT BNP_TYPERESSOURCE,BNP_NATUREPRES, BLO_TYPEARTICLE,BLO_NATUREPIECEG, BLO_SOUCHE,BLO_ARTICLE,BLO_QTEFACT,BLO_DPA,BLO_REMISEPIED,'+
  			 'BLO_DPR,BLO_PUHTDEV,BLO_QTEDUDETAIL,BLO_NUMLIGNE,BLO_N1,BLO_N2,BLO_N3,BLO_N4,BLO_N5, BLO_NATURETRAVAIL, ' +
         'BLO_MONTANTFC, BLO_MONTANTFG, BLO_MONTANTFR, BLO_NUMERO FROM LIGNEOUV '+
  			 'LEFT JOIN ARTICLE     ON GA_ARTICLE=BLO_ARTICLE '+
         'LEFT JOIN NATUREPREST ON BNP_NATUREPRES=GA_NATUREPRES '+
         'WHERE '+
         'AND BLO_NATUREPIECEG = "'+TOBPiece.getValue('GP_NATUREPIECEG')+'" '+
         'AND BLO_SOUCHE="' + TOBPiece.GetValue('GP_SOUCHE') + '" ' +
         'AND BLO_NUMERO=' + IntToStr(TOBPiece.GetValue('GP_NUMERO')) + ' '+
         'AND BLO_INDICEG=' + IntToStr(TOBPiece.GetValue('GP_INDICEG')) + ' '+
         'AND BLO_PUHTDEV <> 0 ' +
         'ORDER BY BLO_NUMLIGNE,BLO_N1,BLO_N2,BLO_N3,BLO_N4,BLO_N5';
  QQ := OpenSql (req,true);
  TOBLocOuvrage.loadDetailDb ('LIGNEOUV','','',QQ,false);
  ferme (QQ);
  ConstitueOuvragesTB ('BLO', TOBPiece,TOBOuvrage,TOBLocOuvrage);
  TOBLocOuvrage.free;
end;

procedure ChargelesOuvragesTBPlat (TOBOuvragePlat,TOBPiece : TOB);
var Req : string;
		QQ : TQuery;
begin

	Req := 'SELECT BNP_TYPERESSOURCE,BNP_NATUREPRES, BOP_TYPEARTICLE,BOP_NATUREPIECEG, BOP_SOUCHE,BOP_ARTICLE,BOP_QTEFACT,BOP_DPA,BOP_REMISEPIED,'+
  			 'BOP_DPR,BOP_PUHTDEV,BOP_TOTALHTDEV,BOP_MONTANTHTDEV, BOP_NUMLIGNE,BOP_NUMORDRE,BOP_N1,BOP_N2,BOP_N3,BOP_N4,BOP_N5, BOP_NATURETRAVAIL, ' +
         'BOP_MONTANTFC, BOP_MONTANTFG, BOP_MONTANTFR, BOP_NUMERO FROM LIGNEOUVPLAT '+
  			 'LEFT JOIN ARTICLE     ON GA_ARTICLE=BOP_ARTICLE '+
         'LEFT JOIN NATUREPREST ON BNP_NATUREPRES=GA_NATUREPRES '+
         'WHERE '+
         'AND BOP_NATUREPIECEG = "'+TOBPiece.getValue('GP_NATUREPIECEG')+'" '+
         'AND BOP_SOUCHE="' + TOBPiece.GetValue('GP_SOUCHE') + '" ' +
         'AND BOP_NUMERO=' + IntToStr(TOBPiece.GetValue('GP_NUMERO')) + ' '+
         'AND BOP_INDICEG=' + IntToStr(TOBPiece.GetValue('GP_INDICEG')) + ' '+
         'AND BOP_TOTALHTDEV <> 0 ' +
         'ORDER BY BOP_NUMORDRE';

  QQ := OpenSql (req,true);

  TOBOuvragePlat.loadDetailDb ('LIGNEOUVPLAT','','',QQ,false);

  ferme (QQ);

end;

procedure TraiteLigneTB (TOBTMP,TOBL : TOB);
begin

  //si intervenant sur ligne article pas prise en compte
  if TOBL.GetString('GLC_NATURETRAVAIL') = '001' then exit;

  if (TobL.getString('GLC_NATURETRAVAIL') = '002') Then
  begin
    TOBL.PutValue('BNP_TYPERESSOURCE', 'ST');
    TOBL.PutValue('LIBNATURE', rechDom('AFTTYPERESSOURCE','ST',false));
  end;

  ChargePrevu(TOBL, TOBTMP, 'GL');

end;

Procedure ChargePrevu(TOBL, TOBTMP : TOB; Prefixe : String; Tps_Prevu : Double);
var MontantPa : Double;
    MontantPr : Double;
    MontantPv : Double;
    //
		NaturePiece : string;
    //
    TypeRes   : String;
    NaturePres: String;
    //
    MontantFG : Double;
    MontantFC : Double;
    MontantFR : Double;
    //
    //TotalMtPA : Double;
    //TotalMtPR : Double;
    //TotalMtPV : Double;
    //
    TotalMtFG : Double;
    TotalMtFC : Double;
    TotalMtFR : Double;
    //
    DPA       : Double;
    DPR       : Double;
    PUHT      : Double;
    //
    //QteduDetail: Double;
    //
    NumeroPiece: Integer;
    //
    TypeArticle: String;
    //
    LigneLog   : String;
    //NoLig      : Integer;
Begin

  LigneLog := '';

	NaturePiece := TOBL.GetValue(Prefixe + '_NATUREPIECEG');
  if pos(NaturePiece,'AFF;DAP')>0  Then NaturePiece := 'PBT';
  //
  TypeRes     := TOBL.GetValue('BNP_TYPERESSOURCE');
  NaturePres  := TOBL.GetValue('BNP_NATUREPRES');

  //FV1 : 04/02/2014 - FS#863 - SCETEC : Distinguer les frais au niveau des champs de prévisionnel
  TypeArticle := TOBL.GetValue(Prefixe + '_TYPEARTICLE');
  If TypeArticle <> 'FRA' then
  begin
    //FV1 : Si pas de nature de prestation alors sur Nature 'Fournitures'
    if TypeRes = '' then TypeRes := 'FOU';
    if NaturePres = '' then NaturePres := 'FOURNITURES';
  end
  else
  begin
    //FV1 : pas de nature de prestation alors sur Nature 'Frais'
    if TypeRes = ''    then TypeRes := 'FRA';
    if NaturePres = '' then NaturePres := 'FRAIS';
  end;
  //
  NumeroPiece := StrToInt(TOBL.GetValue(Prefixe + '_NUMERO'));
  //
  DPA         := TOBL.GetValue(Prefixe + '_DPA');
  DPR         := TOBL.GetValue(Prefixe + '_DPR');
  PUHT        := TOBL.GetValue(Prefixe + '_PUHTDEV');
  //

  if (Tps_Prevu = 0) and (prefixe <> 'BLO') then Tps_Prevu := TOBL.GetValue(Prefixe + '_QTEFACT');
  //
  MontantFG   := TOBL.GetValue(Prefixe + '_MONTANTFG');
  MontantFC   := TOBL.GetValue(Prefixe + '_MONTANTFC');
  MontantFR   := TOBL.GetValue(Prefixe + '_MONTANTFR');
  //
  if prefixe = 'BLO' then
  begin
    //
    MontantPA := (Tps_Prevu*TOBL.GetDouble('BLO_DPA')) ;
    MontantPR := (Tps_Prevu*TOBL.GetDouble('BLO_DPR'));
    MontantPV := (Tps_Prevu*TOBL.GetDouble('BLO_PUHTDEV'));
    //
    if TOBL.GetValue('BLO_REMISEPIED') then
    begin
    	MontantPV := MontantPV - (MontantPV * (TOBL.GetDouble('BLO_REMISEPIED')/100.0));
    end;
    //
    if TOBL.GetValue('BLO_QTEFACT') <> 0 then
      MontantFG := (Tps_Prevu/TOBL.GetDouble('BLO_QTEFACT')) * MontantFG
    else
      MontantFG := 0.0;
    if TOBL.GetValue('BLO_QTEFACT') <> 0 then
      MontantFC := (Tps_Prevu/TOBL.GetDouble('BLO_QTEFACT')) * MontantFC
    else
      MontantFC := 0.0;
    if TOBL.GetValue('BLO_QTEFACT') <> 0 then
      MontantFR := (Tps_Prevu/TOBL.GetDouble('BLO_QTEFACT')) * MontantFR
    else
      MontantFR := 0.0;
    //
  end
  else if prefixe = 'BOP' then
  begin
    MontantPA   := TOBL.GetValue(Prefixe + '_QTEFACT')* DPA;
    MontantPR   := TOBL.GetValue(Prefixe + '_QTEFACT')* DPR;
    MontantPV   := TOBL.GetDouble(Prefixe + '_TOTALHTDEV');
  end
  else
  begin
    MontantPA   := TOBL.GetValue(Prefixe + '_QTEFACT')* DPA;
    MontantPR   := TOBL.GetValue(Prefixe + '_QTEFACT')* DPR;
    MontantPV   := TOBL.GetValue(Prefixe + '_TOTALHTDEV');
    //
{ BRL 9/08 : pas d'avancé par nature de pièce autre que PBT
    AvancePa    := TOBL.GetValue(Prefixe + '_QTEPREVAVANC') * DPA;
    AvancePr    := TOBL.GetValue(Prefixe + '_QTEPREVAVANC') * DPR;
    AvancePV    := TOBL.GetValue(Prefixe + '_TOTALHTDEV')   * (TOBL.GetValue(Prefixe+ '_POURCENTAVANC')/100);
}
  end;


  if TOBL.GetValue(Prefixe + '_TYPEARTICLE')='POU' then
  begin
    // Article de type pourcentage
    MontantPA := 0 ;
    MontantPR := 0;
    MontantPV := (Tps_Prevu * PUHT/100);
  end;

  //Ajout FV1 03/06/2013 : Montant prévu dispatch par nature de pièce
  //TotalMtPA := TOBTMP.GetDouble('PREVU_'+NaturePiece+'_PA');
  //TotalMtPA := TotalMTPA + MontantPA;

  //TotalMtPR := TOBTMP.GetDouble('PREVU_'+NaturePiece+'_PR');
  //TotalMtPR := TotalMTPR + MontantPR;

  //TotalMtPV := TOBTMP.GetDouble('PREVU_'+NaturePiece+'_PV');
  //TotalMtPV := TotalMTPV + MontantPV;

// BRL 13/08 : ces 3 montants ne concernent que la prévision de chantier PBT
  if NaturePiece = 'PBT' then
  begin
    //Ajout FV1 03/06/2013 : Montant prévu dispatch par nature de pièce
    TotalMtFG := TOBTMP.GetValue('PREVU_MONTANTFG');
    TotalMtFG := TotalMTFG + MontantFG;

    TotalMtFC := TOBTMP.GetValue('PREVU_MONTANTFC');
    TotalMtFC := TotalMTFC + MontantFC;

    TotalMtFR := TOBTMP.GetValue('PREVU_MONTANTFR');
    TotalMtFR := TotalMTFR + MontantFR;

    // Ajout BRL 7/11/2012 : Montants des frais détaillés chantier, généraux et répartis
    if TOBTMP.FieldExists('PREVU_MONTANTFG') then TOBTMP.PutValue('PREVU_MONTANTFG', TotalMtFG);
    if TOBTMP.FieldExists('PREVU_MONTANTFC') then TOBTMP.PutValue('PREVU_MONTANTFC', TotalMtFC);
    if TOBTMP.FieldExists('PREVU_MONTANTFR') then TOBTMP.PutValue('PREVU_MONTANTFR', TotalMtFR);
  end;

  if TOBTMP.FieldExists('PREVU_'+NaturePiece+'_MONTANTFG') then
    TOBTMP.PutValue('PREVU_'+NaturePiece+'_MONTANTFG', TOBTMP.GetDouble('PREVU_'+NaturePiece+'_MONTANTFG')+MontantFG);
  if TOBTMP.FieldExists('PREVU_'+NaturePiece+'_MONTANTFC') then
    TOBTMP.PutValue('PREVU_'+NaturePiece+'_MONTANTFC', TOBTMP.GetDouble('PREVU_'+NaturePiece+'_MONTANTFC')+MontantFC);
  if TOBTMP.FieldExists('PREVU_'+NaturePiece+'_MONTANTFR') then
    TOBTMP.PutValue('PREVU_'+NaturePiece+'_MONTANTFR', TOBTMP.GetDouble('PREVU_'+NaturePiece+'_MONTANTFR')+MontantFR);

  TOBTMP.AddChampSupValeur('NUMEROPIECE', NumeroPiece);

  ChargeCumulPrevuFacture(TypeRes, NaturePiece, TOBTMP,MontantPa,MontantPr,MontantPV,Tps_Prevu);

  if not ((Pos(NaturePiece,'FBT;B00;FBP;FBC;DAC')>0) or (Pos(NaturePiece,'ABT;ABP;ABC')>0) or (NaturePiece = 'FAC') or (NaturePiece = 'AVC')) Then
  begin
    if (TypeRes <> 'FOU') and (TOBTMP.GetValue('CEclatNatPrest') = 'X') then
    begin
      Charge_Repartition_Eclatement(NaturePiece + '_' + TypeRes + '_' + NaturePres,MontantPA, MontantPR, MontantPV, TOBTMP)
    end;
  end;
end;

procedure TraiteLigneDetailOuvrageTB (Prefixe : String; TOBTMP,TOBL : TOB;Qte,QteDuDetail : double);
var TpsPrevu : double;
begin
  TpsPrevu := 0;

  If prefixe = 'BLO' then TpsPrevu := Qte/QteDudetail;

  ChargePrevu(TOBL, TOBTMP, Prefixe, TpsPrevu);
end;

procedure TraiteDetailOuvrageTBPlat (TOBTMP,TOBOUV : TOB);
begin
    if (TOBOUV.getString('BOP_NATURETRAVAIL') = '002') Then
    begin
      TOBOUV.PutValue('BNP_TYPERESSOURCE','ST');
      TOBOUV.PutValue('LIBNATURE', rechDom('AFTTYPERESSOURCE','ST',false));
    end;

    TraiteLigneDetailOuvrageTB ('BOP',TOBTMP,TOBOUV,0,0);
end;

procedure TraiteOuvrageTB (TOBTMP,TOBL,TOBOuvrage : TOB);
var IndiceOuv : integer;
		TOBOuv : TOB;
begin
  IndiceOuv := TOBL.GetValue('GL_INDICENOMEN');
  if IndiceOuv =0 then exit;
  if IndiceOuv > TOBOuvrage.detail.count then exit;
  TOBOuv := TOBOuvrage.detail[IndiceOuv-1];
  if TOBOUv = nil then exit;
  TraiteDetailOuvrageTB (TOBTMP,TOBOUV,TOBL.GetDouble('GL_QTEFACT'),1);
end;

procedure TraiteDetailOuvrageTB (TOBTMP,TOBOUV : TOB; Qte,QteDuDetail : double);
var QteSui,QteDuDetailSui : double;
		Indice : integer;
    TOBDet : TOB;
begin

  for indice := 0 to TOBOUV.detail.count -1 do
  begin
  	TOBDet := TOBOUV.detail[Indice];
    if (Tobdet.getString('BLO_NATURETRAVAIL') = '001') Then continue;

    if (Tobdet.getString('BLO_NATURETRAVAIL') = '002') Then
    begin
      TOBDet.PutValue('BNP_TYPERESSOURCE','ST');
      TOBDet.PutValue('LIBNATURE', rechDom('AFTTYPERESSOURCE','ST',false));
    end;

    QteSui := Qte * TOBDet.GetDouble('BLO_QTEFACT');
    if TOBDet.GetDouble('BLO_QTEDUDETAIL') <> 0 then
    begin
       QteDuDetailSui := QteDudetail * TOBDet.GetDouble('BLO_QTEDUDETAIL');
    end
    else
    begin
       QteDuDetailSui := QteDudetail;
    end;

    if TOBDet.detail.count > 0 then
    begin
      TraiteDetailOuvrageTB (TOBTMP,TOBDet,QteSui,QteDuDetailSui);
    end else
    begin
    	TraiteLigneDetailOuvrageTB ('BLO',TOBTMP,TOBdet,QteSui,QteDuDetailSui);
    end;
  end;

end;

procedure TraiteOuvrageTBPlat (TOBTMP,TOBL,TOBOuvragePlat : TOB);
var TOBTT : TOB;
    ii,depart : integer;
begin

  TOBTT := TOBOuvragePlat.findfirst(['BOP_NUMORDRE'],[TOBL.GetInteger('GL_NUMORDRE')],false);

  If TOBTT = nil then exit;
  depart := TOBTT.getIndex;
  for II := depart to TOBOuvragePlat.detail.count -1 do
  begin
    TOBTT := TOBOuvragePlat.detail[II];
    if TOBTT.GetValue('BOP_NUMORDRE')<> TOBL.GetInteger('GL_NUMORDRE') then break;
    TraiteDetailOuvrageTBPLAT (TOBTMP,TOBTT);
  end;
end;



procedure TraiteLaPieceTB (TOBTMP,TOBPiece,TOBOuvrage : TOB);
var Indice : integer;
		TOBL : TOB;
begin

	for indice := 0 to TOBPiece.detail.count -1 do
  begin
    TOBL := TOBPiece.detail[Indice];
    if (Pos (TOBL.GetValue('GL_TYPEARTICLE'),'OUV;OU1;ARP') > 0) and (TOBL.GetValue('GL_INDICENOMEN')>0) then
    begin
      If (TOBL.GetString('GL_NATUREPIECEG') = 'FBT') OR
         //(TOBL.GetString('GL_NATUREPIECEG') = 'B00') OR
         (TOBL.GetString('GL_NATUREPIECEG') = 'FBP') OR
         (TOBL.GetString('GL_NATUREPIECEG') = 'FBC') OR
         (TOBL.GetString('GL_NATUREPIECEG') = 'ABC') OR
         (TOBL.GetString('GL_NATUREPIECEG') = 'ABT') OR
         //(TOBL.GetString('GL_NATUREPIECEG') = 'ABP') OR
         (TOBL.GetString('GL_NATUREPIECEG') = 'DAC') OR
         (TOBL.GetString('GL_NATUREPIECEG') = 'FAC') OR
         (TOBL.GetString('GL_NATUREPIECEG') = 'AVC') then
        TraiteOuvrageTBPlat(TOBTMP,TOBL,TOBOuvrage)
      else
       	TraiteOuvrageTB (TOBTMP,TOBL,TOBOuvrage);
    end else
    begin
    	TraiteLigneTB (TOBTMP,TOBL);
    end;
  end;

end;

procedure DefiniPrevuDetail (TOBTMP,TOBPiece: TOB ; NaturePiece : string; OptionChoixPrevuAvanc: TChoixPrevuAvanc);
var TOBLignes       : TOB;
    TOBOuvrages     : TOB;
begin
	TOBLignes := TOB.Create ('LA PIECE',nil,-1);
  TOBLignes.Dupliquer (TOBPiece,false,true);
  TOBOuvrages := TOB.Create ('LES OUVRAGES',nil,-1);

  ChargelesLignesTB (TOBLIgnes,TOBPiece);
  //
  If  (NaturePiece = 'FBT') OR
      //(NaturePiece = 'B00') OR
      (NaturePiece = 'FBC') OR
      (NaturePiece = 'ABC') OR
      (NaturePiece = 'FBP') OR
      (NaturePiece = 'ABT') OR
      //(NaturePiece = 'ABP') OR
      (NaturePiece = 'DAC') OR
      (NaturePiece = 'FAC') OR
      (NaturePiece = 'AVC') then
    ChargelesOuvragesTBPlat (TOBOuvrages,TOBLignes)
  else
    ChargelesOuvragesTB (TOBOuvrages,TOBLignes);
  //
  TraiteLaPieceTB (TOBTMP,TOBLignes,TOBOuvrages);

  TOBLIgnes.free;
  TOBOuvrages.free;
end;
                                                    
procedure SetPrevuAvanceADetaille (TOBTMP : TOB; NaturePiece : string; OptionChoixPrevuAvanc: TChoixPrevuAvanc; WherePiece : String);
var Req : String;
    TOBPieces   : TOB;
    QQ : TQuery;
    Indice : integer;
    Critere : string;
    RefAffaire : string;
begin

	TOBPieces := TOB.Create ('LES PIECES',nil,-1);

  RefAffaire := 'GP_AFFAIREDEVIS';

  if Naturepiece = 'DBT' then
    Critere := ' AND AFF_ETATAFFAIRE IN ("ACP","TER") '
  else if Naturepiece = 'AFF' then
  BEGIN
    //Critere := ' AND AFF_ETATAFFAIRE IN ("ENC","TER") ';
    //FV1 - 25/06/2015 - FS#1408 - DELABOUDINIERE : en analyse chantier, le prévu contrat n'est pas renseigné.
    Critere := ' AND AFF_ETATAFFAIRE IN ("ACP","ENC","TER") ';
    RefAffaire:='GP_AFFAIRE';
  END
  else if Naturepiece = 'DAP' then
  BEGIN
    Critere := ' AND AFF_ETATAFFAIRE IN ("ACP","TER") ';
  END
  else if NaturePiece = 'ETU' then Critere := ' AND AFF_ETATAFFAIRE="ACP"';

  if WherePiece = '' then
  begin
    if NaturePiece = 'FBT' then
    begin
      //WherePiece := '(GP_NATUREPIECEG IN ("FBT","ABT","FAC","AVC","FBC","ABC") OR (GP_NATUREPIECEG In ("FBP","FPR","DAC","DAP") AND (GP_VIVANTE="X"))) ' +
      WherePiece := '(GP_NATUREPIECEG IN ("FBT","ABT","FAC","AVC","FBC","ABC") OR (GP_NATUREPIECEG In ("FBP","FPR","DAC") AND (GP_VIVANTE="X"))) ' +
                    'AND GP_AFFAIRE="'+TOBTMP.GetValue('BCO_AFFAIRE') + '"';
  	  //WherePiece := ' ((GP_NATUREPIECEG="'+NaturePiece+'") OR (GP_NATUREPIECEG="B00") OR ((GP_NATUREPIECEG="FBP") AND (GP_VIVANTE="X"))) '+
      //'AND GP_AFFAIRE="'+TOBTMP.GetValue('BCO_AFFAIRE') + '"';
    end else
    begin
      WherePiece := ' GP_NATUREPIECEG="'+NaturePiece+'" '+'AND GP_AFFAIRE="'+TOBTMP.GetValue('BCO_AFFAIRE') + '"';
    end;
  end;

	Req := 'SELECT GP_NATUREPIECEG,GP_SOUCHE,GP_NUMERO,GP_INDICEG, GP_MONTANTPA, GP_MONTANTPR, GP_TOTALHTDEV FROM PIECE '+
  			 'LEFT JOIN AFFAIRE ON AFF_AFFAIRE='+RefAffaire+' WHERE '+
  			 WherePiece +
         Critere;

  QQ := OPenSql (req,true);
  TOBPieces.LoadDetailDB ('PIECE','','',QQ,false);
  Ferme (QQ);

  for indice := 0 to TOBPieces.detail.count -1 do
  begin
    DefiniPrevuDetail (TOBTMP,TOBPIeces.detail[Indice],NaturePiece,OptionChoixPrevuAvanc);

    if TOBTMP.FieldExists('PREVU_'+NaturePiece+'_PA') then TOBTMP.PutValue('PREVU_' + NaturePiece +'_PA', TOBTMP.GetValue('PREVU_' + NaturePiece +'_PA') + TOBPIeces.detail[Indice].GetValue('GP_MONTANTPA'));
    if TOBTMP.FieldExists('PREVU_'+NaturePiece+'_PR') then TOBTMP.PutValue('PREVU_' + NaturePiece +'_PR', TOBTMP.GetValue('PREVU_' + NaturePiece +'_PR') + TOBPIeces.detail[Indice].GetValue('GP_MONTANTPR'));
    if TOBTMP.FieldExists('PREVU_'+NaturePiece+'_PV') then TOBTMP.PutValue('PREVU_' + NaturePiece +'_PV', TOBTMP.GetValue('PREVU_' + NaturePiece +'_PV') + TOBPIeces.detail[Indice].GetValue('GP_TOTALHTDEV'));

  end;

  TOBPIeces.free;
  
end;

procedure SetprevuAvance (TOBTMP : TOB; NaturePiece : string ; OptionChoixPrevuAvanc : TChoixPrevuAvanc; DateMvtFin : TdateTime; WherePiece : string='');
begin

  //Modif FV : Dev. prioritaire DSL le 05/06/2012
  CoefFG_param  := GetParamSocSecur('SO_COEFFG', 0);
  Tauxhoraire   := GetParamSocSecur('SO_TAUXHORAIRE', 0);

	if NaturePiece = 'PBT' then
  begin
    if VH_GC.BTCODESPECIF = '002' then
    begin
      SetPrevuAvancePBTVerdon (TOBTMP,OptionChoixPrevuAvanc,DateMvtFin)
    end else
    begin
      SetPrevuAvancePBT (TOBTMP,OptionChoixPrevuAvanc)
    end;
  end else
  begin
    SetPrevuAvanceADetaille (TOBTMP,NaturePiece,OptionChoixPrevuAvanc,WherePiece);
  end;

end;

end.
