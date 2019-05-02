unit UTOFBTArticle_mul;

interface
uses  windows,StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,TiersUtil,
{$IFDEF EAGLCLIENT}
      UtileAGL,eFiche,eFichList,maineagl,Spin,eMul,
{$ELSE}
      DBCtrls,db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FichList,Fiche,Fe_Main,MAJTable,HDB,AglIsoflex,
      mul,
      DBGrids,
{$ENDIF}
      Ent1,
      HCtrls,
      HEnt1,
      HMsgBox,
      UTOF,
      HDimension,
      Dicobtp,
      EntGC,
      HTB97,
      menus,
      ExtCtrls,
      Utob,
      AglInit,
      BTMUL_TOF,
      paramsoc,
      UtilPgi;
Type
     TOF_BTARTICLE_MUL = Class (TOF_BTMUL)
     ButInsert    : Ttoolbarbutton97;
     ButStop      : Ttoolbarbutton97;
     BDuplication : Ttoolbarbutton97;
     FamilleN1,FamilleN2,FamilleN3 : THvalComboBox;
     FamilleN1Ouv,FamilleN2Ouv,FamilleN3Ouv : THvalComboBox;
     FamilleTarif,SousFamilleTarif : THvalComboBox;
     TypeArticle  : THMultiValComboBox;
     TypeArt      : THValComboBox;
     CodeDepot    : THValComboBox;
     GA_BU        : THValComboBox;
     TGA_BU       : THLabel;
        procedure OnUpdate ; override ;
        procedure OnLoad ; override ;
        procedure OnArgument (st : String ) ; override ;
        function GetPlusTypeArticle (Critere : string) : string;
     private
        SType : string;
    		MultiSelect : boolean;
        ffirst : boolean;
        procedure InsertClick(Sender: Tobject);
        procedure StopClick(Sender: Tobject);
        procedure ConstituePlusTypeRessource(TheTypes: string);
        procedure ConstituePlusTypeArticle(TheTypes: string);
        procedure BOuvrirClick(Sender: TObject);
        procedure RenommageEnteteColonnes;
        procedure TypeArticleChange (Sender : Tobject);
        procedure Famille1Change (Sender : TOBject);
        procedure Famille2Change (Sender : TOBject);
        procedure Famille1OuvChange (Sender : TOBject);
        procedure Famille2OuvChange (Sender : TOBject);
        procedure FamilleTarifChange (Sender : TOBject);
        procedure ControleChamp(Champ, Valeur: String);
        procedure DuplicationArticle(Sender: TObject);
     END ;
VAR
   IndiceArt  : Integer;
   XXWhere    : String;
implementation

uses Grids,UFonctionsCBP, TntStdCtrls;

procedure TOF_BTARTICLE_MUL.OnArgument(st : String );
var i           : Integer;
    Deb         : Integer;
    Fin         : Integer;
    X           : integer;
    MenuPop     : Tpopupmenu;
    DroitCreat  : boolean ;
    Critere     : string;
    Chaine      : String;
    Valeur      : String;
    ChampMul    : String;
    ValMul      : string;
    TT   : string;
BEGIN
Inherited;

	ffirst := True;

  FamilleN1     := THValComboBox (ecran.FindComponent ('GA_FAMILLENIV1'));
  FamilleN1Ouv  := THValComboBox (ecran.FindComponent ('GA_FAMILLENIV1_'));
  FamilleN2     := THValComboBox (ecran.FindComponent ('GA_FAMILLENIV2'));
  FamilleN2Ouv  := THValComboBox (ecran.FindComponent ('GA_FAMILLENIV2_'));
  FamilleN3     := THValComboBox (ecran.FindComponent ('GA_FAMILLENIV3'));
  FamilleN3Ouv  := THValComboBox (ecran.FindComponent ('GA_FAMILLENIV3_'));
  //
  if (ecran.Name = 'BTARTICLE_RECH') Or (ecran.Name = 'BTPREST_RECH') then
  Begin
    TypeArt       := THValComboBox(ecran.FindComponent ('GA_TYPEARTICLE'));
    TypeArt.OnChange := TypeArticleChange;
    //
    // Modified by f.vautrain 27/02/2019 09:34:47 / FS#3529 - VERDON : Ajout de la BU dans la fiche article
    if VH_GC.BTCODESPECIF = '002' then
    Begin
      GA_BU   := THValComboBox(Ecran.FindComponent('GA_BU'));
      TGA_BU  := THLabel(Ecran.FindComponent('TGA_BU'));
      if TGA_BU <> nil  then TGA_BU.Visible := True;
      if GA_BU  <> Nil  then GA_BU.Visible  := True;
    end;
    //
  end
  else
  begin
    TypeArticle   := THMultiValComboBox(ecran.FindComponent ('GA_TYPEARTICLE'));
    TypeArticle.OnChange := TypeArticleChange;
  end;
  
  if FamilleN1 <> nil then
  begin
    FamilleN1.OnChange := Famille1Change;
  end;
  if FamilleN2 <> nil then
  begin
    FamilleN2.OnChange := Famille2Change;
  end;
  if FamilleN1Ouv <> nil then
  begin
    FamilleN1Ouv.OnChange := Famille1OuvChange;
  end;
  if FamilleN2Ouv <> nil then
  begin
    FamilleN2Ouv.OnChange := Famille2OuvChange;
  end;

// Ajout BRL : Sélection sur tarif article et sous-famille tarif
  FamilleTarif := THValComboBox (ecran.FindComponent ('GA_TARIFARTICLE'));
  if FamilleTarif <> nil then
  begin
    	FamilleTarif.OnChange := FamilleTarifChange;
    	if (St <> 'MAR') and (St <> 'ARP') then FamilleTarif.enabled:=False;
  end;

  SousFamilleTarif := THValComboBox (ecran.FindComponent ('GA2_SOUSFAMTARART'));
  if SousFamilleTarif <> nil then SousFamilleTarif.enabled := False;

  BDuplication := TToolbarButton97(ecran.FindComponent ('B_DUPLICATION'));
  if BDuplication <> nil then BDuplication.OnClick := DuplicationArticle;

  CodeDepot := THValComboBox(Ecran.FindComponent('GQ_DEPOT'));

  if (ecran.Name = 'BTPREST_RECH') then
  Begin
    if TypeArt <> nil then TypeArt.Visible := False;
  end
  else if copy(ecran.Name,1,16) = 'BTPRESTATION_MUL' then
  begin
    SType := St;
  end
  else if copy(ecran.Name,1,13) = 'BTARTICLE_MUL' then
  begin
    // Modified by f.vautrain 27/02/2019 09:34:47 / FS#3529 - VERDON : Ajout de la BU dans la fiche article
    if VH_GC.BTCODESPECIF = '002' then
    Begin
      GA_BU   := THValComboBox(Ecran.FindComponent('GA_BU'));
      TGA_BU  := THLabel(Ecran.FindComponent('TGA_BU'));
      if TGA_BU <> nil  then TGA_BU.Visible := True;
      if GA_BU  <> Nil  then GA_BU.Visible  := True;
    end;
    //
    SType := St;
    If (st ='PRE') Then
    begin
      TFMul(ecran).SetDBListe ('BTPRESTATIONSMUL');
      Ecran.Caption :='Prestations';
      THEdit(getcontrol('XX_WHERE')).Text := '';
      if TypeArticle <> nil then TypeArticle.Enabled := false;
      SetControlVisible ('GA_LIBREART1',False);
      SetControlVisible ('GA_LIBREART2',False);
      SetControlVisible ('GA_LIBREART3',False);
      SetControlVisible ('TGA_LIBREART1',False);
      SetControlVisible ('TGA_LIBREART2',False);
      SetControlVisible ('TGA_LIBREART3',False);
    end
    else
    If (st ='NOM') Then
    begin
      Critere := 'AND (GA_TYPEARTICLE <> "MAR") AND (GA_TYPEARTICLE <> "POU") AND (GA_TYPEARTICLE <> "ARP")';
      Critere := Critere + 'AND (GA_TYPEARTICLE NOT LIKE "F%")';
      Critere := Critere + 'AND (GA_TYPEARTICLE NOT LIKE "PR%")' ;
      Critere := Critere + 'AND (GA_TYPEARTICLE NOT LIKE "PA%")' ;
      ButInsert := TToolbarButton97 (Getcontrol('Binsert'));
      MenuPop := TpopupMenu(GetControl ('POPCREATNOM'));
      ButInsert.DropdownMenu := MenuPop;
      ButInsert.Width := 35;
      ButInsert.DropDownAlways := true;
      Ecran.Caption :='Ouvrages';
      THEdit(getcontrol('XX_WHERE')).Text := critere;
      //
      if TypeArticle <> nil then TypeArticle.text := '';
      if TypeArticle <> nil then TypeArticle.Plus := 'AND (CO_CODE="NOM" OR CO_CODE="OUV")';
      if TypeArticle <> nil then TypeArticle.Visible := false;
      THlabel(getcontrol('TGA_TYPEARTICLE')).Visible := false;
      //
      SetControlVisible ('GA_LIBREART1',False);
      SetControlVisible ('GA_LIBREART2',False);
      SetControlVisible ('GA_LIBREART3',False);
      SetControlVisible ('TGA_LIBREART1',False);
      SetControlVisible ('TGA_LIBREART2',False);
      SetControlVisible ('TGA_LIBREART3',False);
    end
    else If (st ='FRA') Then
    begin
      Ecran.Caption :='Frais';
      THEdit(getcontrol('XX_WHERE')).Text := '';
      THlabel(getcontrol('TGA_TYPEARTICLE')).Visible := false;
      if TypeArticle <> nil then TypeArticle.Visible := false;
      SetControlVisible ('GA_LIBREART1',False);
      SetControlVisible ('GA_LIBREART2',False);
      SetControlVisible ('GA_LIBREART3',False);
      SetControlVisible ('TGA_LIBREART1',False);
      SetControlVisible ('TGA_LIBREART2',False);
      SetControlVisible ('TGA_LIBREART3',False);
      end
    else If (st ='MAR') Then
    begin
      Critere := 'AND (GA_TYPEARTICLE <> "NOM") AND (GA_TYPEARTICLE <> "OUV")';
      Critere := Critere + 'AND (GA_TYPEARTICLE NOT LIKE "F%")';
      Critere := Critere + 'AND (GA_TYPEARTICLE NOT LIKE "PR%")' ;
      Critere := Critere + 'AND (GA_TYPEARTICLE NOT LIKE "PA%")' ;
      if TypeArticle <> nil then TypeArticle.text := '';
      if TypeArticle <> nil then TypeArticle.Plus := 'AND (CO_CODE="MAR" OR CO_CODE="POU" OR CO_CODE="ARP")';
      THEdit(getcontrol('XX_WHERE')).Text := Critere;
      Ecran.Caption :='Articles';
      ButInsert := TToolbarButton97 (Getcontrol('Binsert'));
      MenuPop := TpopupMenu(GetControl ('POPCREATART'));
      ButInsert.DropdownMenu := MenuPop;
      ButInsert.DropDownAlways := true;
      ButInsert.Width := 35;
    end;
    //Modif BTP
    if (st <> 'NOM') and (st <> 'MAR') and (st <> 'OUV') and (st <>'RG') then
    begin
      ButInsert := TToolbarButton97 (ecran.findcomponent('Binsert'));
      butInsert.OnClick := InsertClick;
    end;
  end
  else
  begin
    if pos('MODIFCODE',st) > 0 then
    Begin
      // Changement de codification article
      {$IFDEF EAGLCLIENT}
      THGrid(GetControl('Fliste')).MultiSelect  := true;
      {$ELSE}
      THDBGrid(GetControl('Fliste')).multiselection := true;
      {$ENDIF}
      TRadioButton (GetControl('MEPPREFIX')).Visible := true;
      TRadioButton (GetControl('CHGCODE')).Checked := true;
      TRadioButton (GetControl('CHGCODE')).Visible := true;
      TToolBarButton97(GetControl('Binsert')).visible := false;
      TToolBarButton97(GetControl('BOuvrir')).visible := false;
      TToolBarButton97(GetControl('BOuvrir1')).visible := true;
      TToolBarButton97(GetControl('BAnnuler1')).visible := false;
      TToolBarButton97(GetControl('BImprimer')).visible := false;
      TToolBarButton97(GetControl('BVoirArticle')).visible := false;
      TToolBarButton97(GetControl('BselectAll')).visible := false;
    End
    else if copy(ecran.Name,1,12) = 'BTPREST_RECH' then
    begin
      repeat
        Critere := uppercase(Trim(ReadTokenSt(St)));
        if Critere <> '' then
        begin
          x := pos('=', Critere);
          if x <> 0 then
          begin
            ChampMul := copy(Critere, 1, x - 1);
            ValMul := copy(Critere, x + 1, length(Critere));
            if ChampMul = 'GA_TYPEARTICLE' then ConstituePlusTypeArticle (ValMul)
            else
            if ChampMul = 'TYPERESSOURCE'  then ConstituePlusTypeRessource (ValMul);
          end;
        end;
      until Critere = '';
      ButInsert := TToolbarButton97 (ecran.findcomponent('Binsert'));
      butInsert.OnClick := InsertClick;
    end
    else if (copy(ecran.Name,1,14) = 'BTARTICLE_RECH') then
    begin
      TT := st;
      repeat
        Critere := uppercase(Trim(ReadTokenSt(TT)));
        if Critere <> '' then
        begin
          x := pos('=', Critere);
          if x <> 0 then
          begin
            ChampMul := copy(Critere, 1, x - 1);
            ValMul := copy(Critere, x + 1, length(Critere));
            if ChampMul = 'STATUTART'  then
            begin
              THMultiValComboBox (getControl('GA_STATUTART')).Text := Stringreplace(ValMul,',',';',[rfReplaceAll]);
            end;
          end;
        end;
      until Critere = '';

      //
      Critere := THEdit(getcontrol('XX_WHERE')).Text;
      XXWhere := THEdit(getcontrol('XX_WHERE')).Text;
      //
      Critere := GetPlusTypeArticle (Critere);
      if critere <> '' then
      begin
        if TypeArt <> nil then TypeArt.Plus  := Critere;
      end;
      //
      ButInsert := TToolbarButton97 (Getcontrol('Binsert'));
      MenuPop := TpopupMenu(GetControl ('POPTYPA'));
      ButInsert.DropdownMenu := MenuPop;
      ButInsert.Width := 35;
      ButInsert.DropDownAlways := true;
      // MODIF LS MULTI-SELECTION
      if TypeArt <> nil then TypeArt.Enabled := (pos('FIXEDTYPEART',st)=0);
      MultiSelect := (Pos('MULTISELECTION',st) > 0);
      SetControlVisible('BSELECTALL', MultiSelect);
      {$IFNDEF EAGLCLIENT}
      TFMul(Ecran).FListe.MultiSelection := MultiSelect;
      {$ELSE}
      TFMul(Ecran).Fliste.MultiSelect := MultiSelect;
      {$ENDIF}
      TToolBarButton97(GetControl('BOUVRIR')).OnClick := BOuvrirClick;
      // --
      if THEdit(ecran.FindComponent ('XX_WHERE1')) <> Nil then
      Begin
        Critere:=THEdit(getcontrol('XX_WHERE1')).Text;
        if critere <> '' then
        Begin
          deb:=pos('%',Critere);
          Critere:=Copy (Critere,deb+1,length(critere)-(Deb+1));
          fin:=pos('%',Critere);
          Ecran.Caption := 'Recherche de : '+copy(critere,1,fin-1);
          TToolBarButton97(GetControl('BANNULER1')).Visible := True;
          ButStop := TToolbarButton97 (ecran.findcomponent('BAnnuler1'));
          ButStop.OnClick := StopClick;
        End;
      end;
      /////
      Critere := St;
      While (Critere <> '') do
      BEGIN
        i:=pos(':',Critere);
        if i = 0 then i:=pos('=',Critere);
        if i <> 0 then
           begin
           Chaine:=copy(Critere,1,i-1);
           Valeur:=Copy (Critere,i+1,length(Critere)-i);
           end
        else
           Chaine := Critere;
        Controlechamp(Chaine, Valeur);
        Critere:=(Trim(ReadTokenSt(St)));
      END;
      //////




      if BDuplication <> nil then BDuplication.Visible := false;
    end;
  END;

  DroitCreat:=ExJaiLeDroitConcept(TConcept(gcArtCreat),False) ;
  if not DroitCreat then
  begin
    if ButInsert <> nil then ButInsert.visible := false;
    if BDuplication <> nil then BDuplication.Visible := false;
  end;

  //uniquement en line
  {*
  SetControlVisible('PCOMPLEMENT',false);
  TFMUL(Ecran).SetDbliste('BTMULARTICLE_S1');
  *}

  Updatecaption(Ecran);
END;

procedure TOF_BTARTICLE_MUL.ControleChamp(Champ, Valeur: String);
begin

  if champ ='GA_TENUESTOCK' then
  begin
      if Valeur = 'X' then
        SetControlProperty('GA_TENUESTOCK', 'Checked', True)
      else
        SetControlProperty('GA_TENUESTOCK', 'Checked', False)
  end;

  if champ ='GA_FERME'      then
  begin
      if Valeur = 'X' then
        SetControlProperty('GA_FERME', 'Checked', True)
      else
        SetControlProperty('GA_FERME', 'Checked', False)
  end;


end;

procedure TOF_BTARTICLE_MUL.OnLoad ;
var st : string;
begin
	inherited;
  if GetparamSocsecur('SO_BTRECHARTAUTO',False) then
  begin
		ffirst := false;
  end;
  if ffirst then
  begin
    ffirst := False;
  end else
  begin
    Thedit(GetControl('XX_WHERE2')).text := '';
  end;
  if FamilleN1 <> nil then
  begin
    if (FamilleN1.Value <> '') then
    begin
      FamilleN2.Plus := 'and CC_LIBRE like "'+FamilleN1.Value+'%"';
      st := st+FamilleN1.Value;
      if FamilleN2.Value <> '' then st := st+FamilleN2.Value;
      if st <> '' then FamilleN3.Plus := 'and CC_LIBRE like "'+st+'%"';
    end;
  end;

RenommageEnteteColonnes;
end;

// MODIF LS MULTI-SELECT

procedure TOF_BTARTICLE_MUL.BOuvrirClick(Sender: TObject);
var iInd : integer;
    FieldArticle : string;
    TobRetour,TobRF : TOB;
begin
  if GetCheckBoxState('RETOUR_CODEARTICLE') = cbChecked then
    FieldArticle := 'GA_CODEARTICLE'
  else
    FieldArticle := 'GA_ARTICLE';

  with TFMul(Ecran) do
  begin
    if not MultiSelect then
    begin
      Retour := Q.FindField(FieldArticle).AsString;
    end else
    begin
      TobRetour := TOB.Create('',nil,-1);
      if not FListe.AllSelected then
      begin
        if FListe.NbSelected = 0 then
        begin
          TobRF := TOB.Create('',TobRetour,-1);
          TobRF.AddChampSupValeur('ARTICLE', Q.FindField(FieldArticle).AsString);
        end else
        begin
          for iInd := 0 to FListe.NbSelected -1 do
          begin
            FListe.GotoLeBookMark(iInd);
    {$IFDEF EAGLCLIENT}
            Q.TQ.Seek (FListe.Row-1) ;
    {$ENDIF}
            TobRF := TOB.Create('',TobRetour,-1);
            TobRF.AddChampSupValeur('ARTICLE',Q.FindField(FieldArticle).AsString);
          end;
        end;
      end
      else
      begin
        Q.First;
        While not Q.Eof do
        begin
          TobRF := TOB.Create('',TobRetour,-1);
          TobRF.AddChampSupValeur('ARTICLE',Q.FindField(FieldArticle).AsString);
          Q.Next;
        end;
      end;
      TheTob := TobRetour;
      if TheTob.Detail.Count > 0 then Retour := TheTob.Detail[0].GetValue('ARTICLE');
    end;
    Close;
  end;

end;
//
procedure TOF_BTARTICLE_MUL.StopClick ;
begin
TFMul(Ecran).Retour:='-1';
Ecran.Close ;
end;

procedure TOF_BTARTICLE_MUL.InsertClick (Sender : Tobject);
var F : TFMul ;
    TArticle : String;
begin
  //
  if not (Ecran is TFMul) then exit;
  //
  F:=TFMul(Ecran) ;
  //
  if (F.Name = 'BTARTICLE_RECH') OR (F.Name = 'BTPREST_RECH') then
    TArticle := TypeArt.Value
  else
    TArticle := TypeArticle.Text;
  //
  if Pos(TArticle,'RES')=0 then
    V_PGI.DispatchTT(7,taCreat,'','','GA_STATUTART=GEN;TYPEARTICLE='+TArticle)
  else
    V_PGI.DispatchTT(6,taCreat,'','','TYPEARTICLE=' + TArticle);
  //
  Ttoolbarbutton97(TFMul(F).FindComponent('Bcherche')).Click  ;
end;

procedure TOF_BTARTICLE_MUL.OnUpdate;
var iCol:integer;
begin
inherited ;
  if not (Ecran is TFMul) then exit;

  if copy(ecran.name,1,13) <> 'BTARTICLE_MUL' then
  begin
    //Mv := THValComboBox (ecran.FindComponent ('GA_TYPEARTICLE'));
    for iCol := 0 to TypeArt.items.Count-1 do
    begin
      if TypeArt.Items.Strings[Icol] = 'Marchandise' then TypeArt.Items.Strings[Icol] := 'Article';
    end;
  end;

  RenommageEnteteColonnes;

end;


function TOF_BTARTICLE_MUL.GetPlusTypeArticle(Critere: string): string;
var Position : integer;
    chaine,ChaineTmp : string;
    first : boolean;
begin

  first := true;

  repeat
    Position := pos ('GA_TYPEARTICLE="',Critere);
    if Position > 0 then
    begin
      chaineTmp := copy (critere,Position+16,3); // bicose longueur des code de tablettes a 3
      if chaineTmp <> '' then
      begin
        if First then
          Chaine := ' AND (CO_CODE IN ("' + ChaineTmp + '"'
        Else
          Chaine := chaine + ', "' + ChaineTmp + '"';
        If First then First := false;
      end;
      critere := copy (critere,Position+17,255);
    end;
  until Position <= 0;
  if chaine  <> '' then Chaine := Chaine + '))';
  result := Chaine;
end;

procedure TOF_BTARTICLE_MUL.ConstituePlusTypeRessource (TheTypes : string);
var lesTypes,unType,TheSelection,TheWhere : string;
begin
  if TheTypes = '' then exit;
  LesTypes := TheTypes;
  TheSelection := '';
  TheWhere := '';
  Untype := READTOKENPipe (LesTypes,',');
  repeat
    if copy(UnType,1,1)='-' then
    begin
      if TheSelection = '' then
        TheSelection := ' AND (CC_CODE NOT IN ("' + copy(UnType,2,255) + '"'
      else
        TheSelection := TheSelection + ',"' + copy(UnType,2,255) + '"';
      //
      if TheWhere = '' then
        TheWhere := '(BNP_TYPERESSOURCE NOT IN ("' + copy(UnType,2,255) + '"'
      else
        TheWhere := TheWhere + ', "' + copy(UnType,2,255) + '"';
    end
    else
    begin
      if TheSelection = '' then
        TheSelection := ' AND (CC_CODE IN ("' + UnType + '"'
      else
        TheSelection := TheSelection + ',"' + UnType + '"';

      if TheWhere = '' then
      begin
        TheWhere := '(BNP_TYPERESSOURCE IN ("' + UnType + '"';
      end else
      begin
        TheWhere := TheWhere + ', "' + UnType + '"';
      end;
    end;
    Untype := READTOKENPipe (LesTypes,',');
  until UnType = '';

  if TheSelection <> '' then
  begin
    TheSelection := TheSelection + '))';
    THValComboBox(GetControl('BNP_TYPERESSOURCE')).plus := TheSelection;
  end;
  if TheWhere <> '' then
  begin
    THEdit(GetControl('XX_WHERE')).text := TheWhere + '))';
  end;
end;

procedure TOF_BTARTICLE_MUL.ConstituePlusTypeArticle(TheTypes: string);
var lesTypes,unType,TheSelection,TheWhere,LastType : string;
		NB : integer;
begin

  if TheTypes = '' then exit;

  if (ecran.Name = 'BTARTICLE_RECH') OR (ecran.Name = 'BTPREST_RECH') then
  begin
    If TypeArt=nil      then Exit;
  end
  else
  begin
    If TypeArticle=nil  then Exit;
  end;

  LesTypes      := TheTypes;

  TheSelection  := '';
  TheWhere      := '';

  Untype        := READTOKENPipe (LesTypes,',');

  NB := 0;

  repeat
  	inc (Nb);
    if copy(UnType,1,1)='-' then
    begin
      if TheSelection = '' then
        TheSelection := 'AND (CO_CODE NOT IN ("' + copy(UnType,2,255) + '"'
      else
        TheSelection := TheSelection + ', "' + copy(UnType,2,255) + '"';
      //
      if TheWhere = '' then
        TheWhere := '(GA_TYPEARTICLE NOT IN ("' + copy(UnType,2,255) +'"'
      else
      begin
        TheWhere := TheWhere + ', "' + copy(UnType,2,255) + '"'
      end;
    end
    else
    begin
    	LastType := UnType;
      if TheSelection = '' then
        TheSelection := ' AND (CO_CODE IN ("' + UnType + '"'
      else
        TheSelection := TheSelection + ', "' + UnType + '"';
      //
      if TheWhere = '' then
        TheWhere := '(GA_TYPEARTICLE IN ("' + UnType + '"'
      else
        TheWhere := TheWhere + ', "' + UnType + '"';
    end;
    Untype := READTOKENPipe (LesTypes,',');
  until UnType = '';

  if TheSelection <> '' then
  begin
    TheSelection := TheSelection + '))';
    if (ecran.Name = 'BTARTICLE_RECH') OR (ecran.Name = 'BTPREST_RECH') then
    begin
      TypeArt.plus := TheSelection;
      if (NB = 1) and (LastType <> '')  then
      begin
        TypeArt.Value := LastType;
      end;
    end
    else
    begin
      TypeArticle.plus := TheSelection;
      if (NB = 1) and (LastType <> '')  then
      begin
        TypeArticle.Items[NB-1] := LastType;
      end;
    end;
  end;
  //
  if TheWhere <> '' then
  begin
    THEdit(GetControl('XX_WHERE')).text := TheWhere + '))';
  end;

end;

procedure TOF_BTARTICLE_MUL.RenommageEnteteColonnes;
var i : integer;
{$IFDEF EAGLCLIENT}
		Gr : THgrid;
{$ELSE}
		Gr : THDbgrid;
{$ENDIF}

    stChamp,Libelle : string;
begin
	Gr := TFMul(ecran).fliste;
{$IFDEF EAGLCLIENT}
	For i:=0 to Gr.ColCount -1 do
  Begin
    StChamp := TFMul(Ecran).Q.FormuleQ.GetFormule(Gr.ColNames [i]);
    if copy(UpperCase (stChamp),1,6)='LIBPRE' then
    begin
      libelle := RechDom('GCLIBFAMILLE','LF'+Copy(stChamp,7,1),false);
			TFMul(ecran).SetDisplayLabel (StChamp,TraduireMemoire(Libelle));
    end else if copy(UpperCase (stChamp),1,11)='GA_LIBREART' then
    begin
      libelle := RechDomZoneLibre ('AT'+Copy(stChamp,12,1),false);
			TFMul(ecran).SetDisplayLabel (StChamp,TraduireMemoire(Libelle));
    end else if copy(UpperCase (stChamp),1,12)='GA_DATELIBRE' then
    begin
      libelle := RechDomZoneLibre ('AD'+Copy(stChamp,13,1),false);
			TFMul(ecran).SetDisplayLabel (StChamp,TraduireMemoire(Libelle));
    end else if copy(UpperCase (stChamp),1,11)='GA_VALLIBRE' then
    begin
      libelle := RechDomZoneLibre ('AM'+Copy(stChamp,12,1),false);
			TFMul(ecran).SetDisplayLabel (StChamp,TraduireMemoire(Libelle));
    end else if copy(UpperCase (stChamp),1,12)='GA_CHARLIBRE' then
    begin
      libelle := RechDomZoneLibre ('AC'+Copy(stChamp,13,1),false);
			TFMul(ecran).SetDisplayLabel (StChamp,TraduireMemoire(Libelle));
    end else if copy(UpperCase (stChamp),1,12)='GA_BOOLLIBRE' then
    begin
      libelle := RechDomZoneLibre ('AB'+Copy(stChamp,13,1),false);
			TFMul(ecran).SetDisplayLabel (StChamp,TraduireMemoire(Libelle));
    end;
  end;

{$ELSE}
	For i:=0 to Gr.Columns.Count-1 do
  Begin
    StChamp := TFMul(Ecran).Q.FormuleQ.GetFormule(Gr.Columns[i].FieldName);
    if copy(UpperCase (stChamp),1,6)='LIBPRE' then
    begin
      libelle := RechDom('GCLIBFAMILLE','LF'+Copy(stChamp,7,1),false);
{$IFNDEF AGL581153}
			TFMul(ecran).SetDisplayLabel (StChamp,TraduireMemoire(Libelle));
{$else}
			TFMul(ecran).SetDisplayLabel (i,TraduireMemoire(Libelle));
{$endif}
    end else if copy(UpperCase (stChamp),1,11)='GA_LIBREART' then
    begin
      libelle := RechDomZoneLibre ('AT'+Copy(stChamp,12,1),false);
{$IFNDEF AGL581153}
			TFMul(ecran).SetDisplayLabel (StChamp,TraduireMemoire(Libelle));
{$else}
			TFMul(ecran).SetDisplayLabel (i,TraduireMemoire(Libelle));
{$endif}
    end else if copy(UpperCase (stChamp),1,12)='GA_DATELIBRE' then
    begin
      libelle := RechDomZoneLibre ('AD'+Copy(stChamp,13,1),false);
{$IFNDEF AGL581153}
			TFMul(ecran).SetDisplayLabel (StChamp,TraduireMemoire(Libelle));
{$else}
			TFMul(ecran).SetDisplayLabel (i,TraduireMemoire(Libelle));
{$endif}
    end else if copy(UpperCase (stChamp),1,11)='GA_VALLIBRE' then
    begin
      libelle := RechDomZoneLibre ('AM'+Copy(stChamp,12,1),false);
{$IFNDEF AGL581153}
			TFMul(ecran).SetDisplayLabel (StChamp,TraduireMemoire(Libelle));
{$else}
			TFMul(ecran).SetDisplayLabel (i,TraduireMemoire(Libelle));
{$endif}
    end else if copy(UpperCase (stChamp),1,12)='GA_CHARLIBRE' then
    begin
      libelle := RechDomZoneLibre ('AC'+Copy(stChamp,13,1),false);
{$IFNDEF AGL581153}
			TFMul(ecran).SetDisplayLabel (StChamp,TraduireMemoire(Libelle));
{$else}
			TFMul(ecran).SetDisplayLabel (i,TraduireMemoire(Libelle));
{$endif}
    end else if copy(UpperCase (stChamp),1,12)='GA_BOOLLIBRE' then
    begin
      libelle := RechDomZoneLibre ('AB'+Copy(stChamp,13,1),false);
{$IFNDEF AGL581153}
			TFMul(ecran).SetDisplayLabel (StChamp,TraduireMemoire(Libelle));
{$else}
			TFMul(ecran).SetDisplayLabel (i,TraduireMemoire(Libelle));
{$endif}
    end;
  end;
{$ENDIF}

end;

procedure TOF_BTARTICLE_MUL.TypeArticleChange(Sender: Tobject);
var icol : integer;
    CC : THLabel;
    TArticle : String;
begin

  if (ecran.Name = 'BTARTICLE_RECH') OR (ecran.Name = 'BTPREST_RECH') then
  begin
    TArticle := TypeArt.Value;
    // Modified by f.vautrain 08/01/2019 13:52:05 - FS#3437 - REPORT V10 - en saisie de pièce, la recherche d’article plante sans ouvrir la fenêtre de recherche
    if TArticle = '' then exit;
  end
  else
  begin
    TArticle := TypeArticle.Text;
    TArticle := copy(TArticle,0, length(TArticle)-1);
  end;

  if FamilleTarif <> nil then
  begin
    If pos(TArticle,'MAR;ARP') > 0 then
      FamilleTarif.enabled := True
    else
      FamilleTarif.enabled := False;
  end;

  THEdit(getcontrol('XX_WHERE')).Text := '';

	if Pos(TArticle, 'OUV') > 0 then
  	TFMul(ecran).SetDBListe ('BTMULOUVRAGE')
  else
	if Pos(TArticle, 'PRE') > 0 then
  	TFMul(ecran).SetDBListe ('BTPRESTATIONSMUL')
  else
  begin
  	TFMul(ecran).SetDBListe ('BTRECHARTICLE');
    THEdit(getcontrol('XX_WHERE')).Text := XXWhere;
  end;

  if Assigned(FamilleN1) then FamilleN1.Visible := (Pos(TArticle, 'OUV') = 0);
  if Assigned(FamilleN2) then FamilleN2.Visible := (Pos(TArticle, 'OUV') = 0);
  if Assigned(FamilleN3) then FamilleN3.Visible := (Pos(TArticle, 'OUV') = 0);
  //
  if assigned(FamilleN1Ouv) then
  begin
    //FamilleN1Ouv.value    := '';
    FamilleN1Ouv.Visible  := (Pos(TArticle, 'OUV') > 0);
  end;
  if assigned(FamilleN2Ouv) then
  begin
    //FamilleN2Ouv.value    := '';
    FamilleN2Ouv.Visible  := (Pos(TArticle, 'OUV') > 0);
  end;
  if assigned(FamilleN3Ouv) then
  begin
    //FamilleN3Ouv.value    := '';
    FamilleN3Ouv.Visible  := (Pos(TArticle, 'OUV') > 0);
  end;

  for iCol:=1 to 3 do
  begin
    CC:=THLabel(GetCOntrol('TGA_FAMILLENIV'+InttoStr(iCol)));
    if copy(ecran.name,1,13) <>'BTARTICLE_MUL' then
    begin
      if (Pos(TArticle, 'OUV')>0) then
      begin
        CC.Caption:=RechDom('BTLIBOUVRAGE','BO'+InttoStr(iCol),false);
      end else
      begin
        CC.Caption:=RechDom('GCLIBFAMILLE','LF'+InttoStr(iCol),false);
      end;
    end;
  end;

  if GetparamSocsecur('SO_BTRECHARTAUTO',False) then
  begin
		TToolbarButton97 (GetControl('BCherche')).Click;	
  end;

end;

procedure TOF_BTARTICLE_MUL.Famille1Change(Sender: TOBject);
begin
  if GetParamSoc('SO_GCFAMHIERARCHIQUE')=True then
  begin
    FAMILLEN2.Plus := '';
    FAMILLEN3.Plus := '';
    if (FAMILLEN1.Value <> '') then
    begin
      FAMILLEN2.Plus := 'and CC_LIBRE like "'+FAMILLEN1.Value+'%"';
      FAMILLEN3.Plus := 'and CC_LIBRE like "'+FAMILLEN1.Value+'%"';
    end;
  end;
end;

procedure TOF_BTARTICLE_MUL.Famille2Change(Sender: TOBject);
var st : string;
begin
  if GetParamSoc('SO_GCFAMHIERARCHIQUE')=True then
  begin
    st := '';
    FAMILLEN3.Plus := '';
    if FAMILLEN1.Value <> '' then st := st+FAMILLEN1.Value;
    if FAMILLEN2.Value <> '' then st := st+FAMILLEN2.Value;
    if st <> '' then FAMILLEN3.Plus := 'and CC_LIBRE like "'+st+'%"';
  end;

end;

procedure TOF_BTARTICLE_MUL.Famille1OuvChange(Sender: TOBject);
begin
  if GetParamSoc('SO_GCFAMHIERARCHIQUE')=True then
  begin
    FAMILLEN2Ouv.Plus := '';
    FAMILLEN3Ouv.Plus := '';
    if (FAMILLEN1Ouv.Value <> '') then
    begin
      FAMILLEN2Ouv.Plus := 'and CC_LIBRE like "'+FAMILLEN1Ouv.Value+'%"';
      FAMILLEN3Ouv.Plus := 'and CC_LIBRE like "'+FAMILLEN1Ouv.Value+'%"';
    end;
  end;
end;

procedure TOF_BTARTICLE_MUL.Famille2OuvChange(Sender: TOBject);
var st : string;
begin

  if GetParamSoc('SO_GCFAMHIERARCHIQUE')=True then
  begin
    st := '';
    FAMILLEN3Ouv.Plus := '';
    if FAMILLEN1Ouv.Value <> '' then st := st+FAMILLEN1Ouv.Value;
    if FAMILLEN2Ouv.Value <> '' then st := st+FAMILLEN2Ouv.Value;
    if st <> '' then FAMILLEN3Ouv.Plus := 'and CC_LIBRE like "'+st+'%"';
  end;

end;

procedure TOF_BTARTICLE_MUL.FamilleTarifChange(Sender: TOBject);
begin
    SousFamilleTarif.Plus := '';
    if (FamilleTarif.Value <> '') then
    begin
      SousFamilleTarif.enabled := True;
      SousFamilleTarif.Plus := 'and BSF_FAMILLETARIF like "'+FamilleTarif.Value+'%"';
    end else
    begin
      SousFamilleTarif.enabled := False;
    end;
end;

procedure TOF_BTARTICLE_MUL.DuplicationArticle(Sender : TObject);
var F           : TFMul ;
    FieldArticle: string;
    TArticle : String;
    CodeArticle : string;
begin

  if not (Ecran is TFMul) then exit;

  F:=TFMul(Ecran) ;

  if GetCheckBoxState('RETOUR_CODEARTICLE') = cbChecked then
    FieldArticle := 'GA_CODEARTICLE'
  else
    FieldArticle := 'GA_ARTICLE';

  CodeArticle := F.Q.FindField(FieldArticle).AsString;

  TArticle := TypeArticle.Text;

  if Pos(TArticle,'RES') = 0 then
    V_PGI.DispatchTT(7,taCreat,CodeArticle,'','DUPLICATION='+CodeArticle+';TYPEARTICLE='+TArticle)
  else
    V_PGI.DispatchTT(6,taCreat,CodeArticle,'','DUPLICATION='+CodeArticle);

  //F.Close;

end;


Initialization
registerclasses([TOF_BTARTICLE_MUL]);
end.
