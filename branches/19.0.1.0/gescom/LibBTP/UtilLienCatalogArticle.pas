unit UtilLienCatalogArticle;

interface
uses  variants,
  Windows, Messages,SysUtils,Classes,
  MajTable,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  UtilArticle,
  UTob,
  HEnt1,
  HCtrls;

function TrouveArticleViaCatalogue (CodeProdCatalog : string; var TOBArt,TOBCata : TOB ;
                                    TOBCatalogu,TobArticles: TOB; OptionCat : ToptionCatArticles;
                                    CurrentCatalog : TCatalog; DatePiece : TdateTime; VenteAchat : string;
                                    ActionBase : TactionFiche) : boolean;

implementation
uses FactArticle,
     EntGc,
     FactFormule,
     UFonctionsCBP,
     Ent1,
     FactUtil,
     hmsgbox
     ;

function TrouveArticleViaCatalogue (CodeProdCatalog : string; var TOBArt,TOBCata : TOB ;
                                    TOBCatalogu,TobArticles: TOB; OptionCat : ToptionCatArticles;
                                    CurrentCatalog : TCatalog; DatePiece : TdateTime; VenteAchat : string;
                                    ActionBase : TactionFiche) : boolean;
var QQ : TQuery;
    TOBArtRef : TOB;
    CodeArticle,Article : string;
    MtPaf : double;
    RechArt : T_RechArt;
    ActionArt : TActionFiche;
begin
  Result := false;
  //
  TOBArt := FindTOBArtInCatalog (TOBCatalogu,TOBArticles,UpperCase(CodeProdCatalog),CurrentCatalog);
  if TOBART <> nil then
  begin
    // cas de figure ou le catalogue est déjà rattaché a un article
    Result := true;
  end else
  begin
    TOBCata := CreerTOBCata(TOBCatalogu);
    //
    RechArt := TrouverCatalog(CurrentCatalog.Fournisseur, UpperCase(CodeProdCatalog), DatePiece, TOBCata);
    if RechArt = traAucun then
    begin
      TOBCata.free; TOBCata := nil;
      TOBArt := FindTOBArtSais(TOBArticles, Article);
      if TOBArt = nil then
      begin
        Article := CodeArticleUnique(CodeProdCatalog,'','','','','');
        TOBArt := CreerTOBArt(TOBArticles);
        QQ := OpenSql('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE,"-" AS PRIXFROMCATALOGUE FROM ARTICLE A ' +
                      'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES ' +
                      'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_ARTICLE="' + Article + '"', true, -1, '', True);
        if not QQ.EOF then
        begin
          ChargerTobArt(TOBArt, nil, VenteAchat, '', QQ);
          Result := True;
        end else
        begin
          TOBArt.free; TOBArt := nil;
        end;
        Ferme(QQ);
      end;
    end else
    begin
      if Trim(TOBCata.GetValue('GCA_ARTICLE')) <> '' then
      begin
        Article := TOBCata.GetValue('GCA_ARTICLE');
        TOBArt := FindTOBArtSais(TOBArticles, Article);
        if TOBArt = nil then
        begin
          TOBArt := CreerTOBArt(TOBArticles);
          QQ := OpenSql('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A ' +
                        'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES ' +
                        'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_ARTICLE="' + Article + '"', true, -1, '', True);
          if not QQ.EOF then
          begin
            ChargerTobArt(TOBArt, nil, VenteAchat, '', QQ);
            Result := True;
          end else
          begin
            TOBArt.free; TOBArt := nil;
          end;
          Ferme(QQ);
        end else
        begin
          Result := True;
        end;
      end else
      begin
        if not OptionCat.CreatArtFromCat then
        begin
          if OptionCat.ArtGenerique = '' then
          begin
            Result := false;
          end else
          begin
            TOBArt := FindTOBArtSais(TOBArticles, OptionCat.ArtGenerique);
            if TOBArt = nil then
            begin
              TOBArt := CreerTOBArt(TOBArticles);
              QQ := OpenSql('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE,"X" AS PRIXFROMCATALOGUE FROM ARTICLE A ' +
                            'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES ' +
                            'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_ARTICLE="' + OptionCat.ArtGenerique + '"', true, -1, '', True);
              if not QQ.EOF then
              begin
                ChargerTobArt(TOBArt, nil, VenteAchat, '', QQ);
              end else
              begin
                TOBArt.Free; TOBART := nil;
              end;
              Ferme(QQ);
            end;
            if TOBART <> nil then
            begin
              TOBArt.SetString('PRIXFROMCATALOGUE','X');
              TOBART.SetString('GA_LIBELLE',TOBCata.getString('GCA_LIBELLE'));
              if TOBCata.GetDouble('GCA_PRIXVENTE') <> 0 then
              begin
                MtPaf := TOBCata.GetDouble('GCA_PRIXVENTE');
              end else
              begin
                mtPaf := TOBCata.GetDouble('GCA_DPA');
              end;
              TOBART.SetDouble('GA_PAUA',mtpaf);
              TOBART.SetDouble('GA_PAHT',mtpaf); TOBART.SetDouble('GA_DPA',mtpaf); TOBART.SetDouble('GA_DPR',mtpaf); TOBART.SetDouble('GA_PVHT',mtpaf);

              TOBART.SetString('GA_QUALIFUNITEACH',TOBCata.GetString('GCA_QUALIFUNITEACH'));
              TOBART.SetString('GA_QUALIFUNITEVTE',TOBCata.GetString('GCA_QUALIFUNITEACH'));
              RecalculPrPV (TOBArt,TOBCata);
              result := true;
            end;
          end;
        end else
        begin
          // article non trouvé en bibliothèque et option de création d'article en bibliothèque positionné
          Article := OptionCat.ArtEtalon;
          TOBArtRef := FindTOBArtSais(TOBArticles, Article);
          if TOBArtRef = nil then
          begin
            QQ := OpenSql('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE,"-" AS PRIXFROMCATALOGUE FROM ARTICLE A ' +
                          'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES ' +
                          'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_ARTICLE="' + Article + '"', true, -1, '', True);
            if not QQ.EOF then
            begin
              TOBArtRef := CreerTOBArt(TOBArticles);
              ChargerTobArt(TOBArtRef, nil, VenteAchat, '', QQ);
            end else
            begin
              TOBArt.Free; TOBArt := nil;
            end;
            Ferme(QQ);
          end;
          if TOBArtRef <> nil then
          begin
            TOBArt := CreerTOBArt(TOBArticles);
            TOBArt.Dupliquer(TOBArtRef,false,true);
            CodeArticle := copy(Trim(CurrentCatalog.Prefixe)+Trim(TOBCata.GetString('GCA_REFERENCE')),1,18);
            Article := CodeArticleUnique(CodeArticle,'','','','','');
            TOBArt.SetString('GA_ARTICLE',Article);
            TOBArt.SetString('GA_CODEARTICLE',CodeArticle);
            TOBCata.setString('GCA_ARTICLE',Article);
            TOBART.SetString('GA_LIBELLE',TOBCata.GetString('GCA_LIBELLE'));
            TOBART.SetString('GA_COMMENTAIRE','');
            TOBART.SetString('GA_FOURNPRINC',TOBCata.GetString('GCA_TIERS'));
            if TOBCata.GetDouble('GCA_PRIXVENTE') <> 0 then
            begin
              MtPaf := TOBCata.GetDouble('GCA_PRIXVENTE');
            end else
            begin
              mtPaf := TOBCata.GetDouble('GCA_DPA');
            end;
            TOBART.SetDouble('GA_PAUA',mtpaf);
            TOBART.SetDouble('GA_PAHT',mtpaf); TOBART.SetDouble('GA_DPA',mtpaf); TOBART.SetDouble('GA_DPR',mtpaf); TOBART.SetDouble('GA_PVHT',mtpaf);

            TOBART.SetString('GA_QUALIFUNITEACH',TOBCata.GetString('GCA_QUALIFUNITEACH'));
            TOBART.SetString('GA_QUALIFUNITEVTE',TOBCata.GetString('GCA_QUALIFUNITEACH'));
            RecalculPrPV (TOBArt,TOBCata);

            TRY
              if OptionCat.FamComptaToUse <> '' then TOBART.SetString('GA_COMPTAARTICLE',OptionCat.FamComptaToUse);
              try
                BEGINTRANS;
                TOBART.InsertDB(nil);
                TOBCata.UpdateDB(false);
                // ouverture de la fiche si ok
                if OptionCat.OpenFormAfterCreat then
                begin
                  if ActionBase <> TaConsult then ActionArt := taModif else ActionArt := taConsult;
                  if ((ActionArt <> taConsult) and (not ExJaiLeDroitConcept(TConcept(gcArtModif), false))) then
                    ActionArt := taConsult;
                  ZoomArticle(Article, TOBART.GetValue('GA_TYPEARTICLE'), ActionArt);
                  if ActionArt <> taConsult then ChargerTobArt(TOBArt, nil, VenteAchat, TOBArt.GetString('GA_ARTICLE'), QQ);
//                  TOBArt.UpdateDB(false);
                end;
                Result := True;
                COMMITTRANS;
              Except
                On e:Exception do
                begin
                  ROLLBACK;
                  PgiError(E.Message);
                  Result := false;
                end;
              end;
            FINALLY
              TOBArtRef.Free;
            END;
          end;
        end;
      end;
    end;
  end;
end;

end.
