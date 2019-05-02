unit UtilBSV;

interface
uses
  Classes,
  SysUtils,
  uTob,
  UDefGlobals,
  Hctrls,
  Db, {$IFNDEF DBXPRESS} dbTables {$ELSE} uDbxDataSet {$ENDIF}
  ;

type
  TFUnctionBSV = class (TObject)
    class function GetIDBSVDOC(TOBpiece : TOB) : string;
    class function EncodeRefPresqueCPGescom(TOBPiece: TOB): String; overload;
    class function EncodeRefPresqueCPGescom (Cledoc: r_cledocb): String; overload;
  end;

implementation


class function TFUnctionBSV.GetIDBSVDOC (TOBpiece : TOB) : string;
var RefGEscom : string;
    QQ : Tquery;
begin
  if TOBpiece.getString('GP_BSVREF')<>'' then
  begin
    result := TOBpiece.getString('GP_BSVREF')
  end else
  begin
    RefGescom := EncodeRefPresqueCPGescom(TOBpiece);
    TRY
      QQ := openSql('SELECT BM4_IDZEDOC FROM BASTENT WHERE BM4_REFGESCOM="'+RefGescom+'"',True,1,'',true);
      if not QQ.eof then
      begin
        Result := IntToStr(QQ.Fields[0].asInteger);
      end;
      ferme (QQ);
    EXCEPT
      Result := '';
    END;
  end;
end;

class Function TFUnctionBSV.EncodeRefPresqueCPGescom ( TOBPiece : TOB ) : String ;
BEGIN
  Result:='' ; if TOBPiece=Nil then Exit ;
  Result:=TOBPiece.GetValue('GP_NATUREPIECEG')+';'+TOBPiece.GetValue('GP_SOUCHE')+';'+'%'+';'
         +IntToStr(TOBPiece.GetValue('GP_NUMERO'))+';'+IntToStr(TOBPiece.GetValue('GP_INDICEG'))+';' ;
END ;

class Function TFUnctionBSV.EncodeRefPresqueCPGescom ( Cledoc : r_cledocb ) : String ;
BEGIN
  Result:='' ; if Cledoc.NaturePiece ='' then Exit ;
  Result:=Cledoc.NaturePiece +';'+Cledoc.Souche +';'+'%'+';'
         +IntToStr(Cledoc.NumeroPiece )+';'+IntToStr(Cledoc.Indice )+';' ;
END ;


end.
