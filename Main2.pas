method MainForm.ActualizarCampoEstado(CardCode:String;FrozenFor:String;TipoSocio:integer);
var
    LoadReg     : SAPbobsCOM.BusinessPartners;
    ErrCode     : Integer;
    sErrMsg     : String;
    SErrCode    : string;
begin
    try
        LoadReg  :=  SAPbobsCOM.BusinessPartners(FCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners));
           try
           if TipoSocio=100 then
           begin
              If LoadReg.GetByKey(CardCode) Then
              begin
                 if FrozenFor = 'N' then
	                  LoadReg.UserFields.Fields.Item('U_Estado').Value := 0 
                 else
	                  LoadReg.UserFields.Fields.Item('U_Estado').Value := 1; 
                 ErrCode := LoadReg.Update();
                 if ErrCode<>0 then
                 Begin                
                    SErrCode:=FCompany.GetLastErrorDescription;
                    AddLog('Error al actualizar el Estado del socio '+SErrCode+' - Codigo '+ErrCode.ToString());                
                    FCompany.GetLastError(out ErrCode,out sErrMsg);                
                 End;
                 //Else
                 //begin
                   //AddLog(String.Format('Actualización de CUPOS sobre SAP, registro : {0} ',FCompany.GetNewObjectKey)); 
                 //End;
              end; 
           end;
           finally
             SBO_f._ReleaseCOMObject(LoadReg);
           end;
    except
       on e: exception do
       AddLog('Error  : ' + e.Message + ' ** Trace: ' + e.StackTrace);
    end;   
end;


method MainForm.ExisteSocio(Rut:string;CardCode:String;FrozenFor:string;TipoSocio:integer): Boolean;
var
   QueryExi    : OdbcCommand;
   readerExi   : OdbcDataReader;
begin
  try
    result:=false;
    QueryExi := new OdbcCommand("SELECT ACCESO_RUT
                                 FROM acceso_llacolen
                                 WHERE ACCESO_RUT = substring('" + Rut + "',1,length('" + Rut + "')-2)", Conn);
    readerExi := QueryExi.ExecuteReader();

    while (readerExi.Read()) do
	  begin
         result:=true;
         break;
    end;

    if result = false then
       //actualizar campo Estado segun el fronzefor si es N(pag) estado 0; si es Y(NoPag) estado 1 **************************
       ActualizarCampoEstado(CardCode,FrozenFor,TipoSocio);

  except
    on e: exception do
    begin
         AddLog('Error al procesar registros : ' + e.Message + ' ** Trace: ' + e.StackTrace);
    end;
  end;
end;


method MainForm.LoadCuposMySQL(Rut:string);
var
    QueryUpd     : OdbcCommand;
begin
   Rut := Rut.Substring(0,Rut.Length-2); 
   try
     QueryUpd := new OdbcCommand("UPDATE acceso_llacolen
                                  SET ACCESO_CUPOS_VISITAS = 20   
                                  WHERE ACCESO_RUT = '" + Rut + "'", Conn);
     QueryUpd.ExecuteNonQuery();
  except
    on e: exception do
    begin
         AddLog('Error al procesar registros : ' + e.Message + ' ** Trace: ' + e.StackTrace);
    end;
  end;
end;

method MainForm.SincronizarSAP();
var
  Rut       : String;
  FechaLoad : string;
  FrozenFor : String;
  TipoSocio : Integer;
  CardCode  : String;
  Estado    : Integer;
begin
    try
        //BUSCO FECHA DE VENCIMIENTO EN TABLA DEFINIDA POR USUARIO, EN LA FECHA DE INICIO PROCESAR LOS PAGADOS Y NO PAGADOS
        FechaLoad:= String.Format("{0:yyyyMMdd}", Datetime.Now);
        oRecordBP.DoQuery(string.Format("SELECT U_Periocidad as Semestre
                                         FROM [@PERIODO]
                                         WHERE {0} = convert(varchar, U_FechaIni, 112)",FechaLoad));
        if Not oRecordBP.EoF then
        begin
             oRecordCU.DoQuery("Select CardCode, FrozenFor, U_Tickets, LicTradNum, U_Estado, GroupCode
                                from OCRD 
                                where CardType = 'C' and (GroupCode=100 or GroupCode=102)");
             While Not oRecordCU.EoF do
             begin
               //CHEQUEAMOS SI ESTA PAGADO O NO; N(ACTIVADO), Y(INACTIVO)
                CardCode := oRecordBP.Fields.Item('CardCode').Value.tostring;
                FrozenFor := oRecordBP.Fields.Item('FrozenFor').Value.tostring;
                TipoSocio := system.Int16(oRecordBP.Fields.Item('GroupCode').Value);
                Estado    := system.Int32(oRecordBP.Fields.Item('U_Estado').Value);
               //SI ES PAGADO, ENTONCES DESDE DICHA FECHA SOCIOS ACTUALIZA CUPOS A 20
               if (FrozenFor = 'N') and (Estado=0) then
               begin
                 Rut := oRecordCU.Fields.Item('LicTradNum').Value.ToString;
                 LoadCuposMySQL(Rut); 
               end
               //SI NO A PAGADO, ENTONCES EN SAP ACTUALIZAMOS EL CAMPO ESTADO A 1
               else if (FrozenFor = 'Y') and (TipoSocio=100) then
               begin
                  ActualizarCampoEstado(CardCode,FrozenFor,TipoSocio);
               end;
             end;
         end 
         else
         begin
             oRecordCU.DoQuery("Select CardCode, FrozenFor, U_Tickets, LicTradNum, U_Estado, GroupCode
                                from OCRD 
                                where CardType = 'C' and 
                                      U_Estado = 1 and 
                                      FrozenFor = 'N' and
                                      (GroupCode=100 or GroupCode=102)");
             While Not oRecordCU.EoF do
             begin
               //SI SOCIO PAGÓ, DESPUÉS DE FECHA DE VENCIMIENTO, SE ACTUALIZA A 20 CUPOS Y ESTADO = 0
                CardCode := oRecordBP.Fields.Item('CardCode').Value.tostring;
                FrozenFor := oRecordBP.Fields.Item('FrozenFor').Value.tostring;
                TipoSocio := system.Int16(oRecordBP.Fields.Item('GroupCode').Value);
                Rut := oRecordCU.Fields.Item('LicTradNum').Value.ToString;
                LoadCuposMySQL(Rut); 
                ActualizarCampoEstado(CardCode,FrozenFor,TipoSocio);
             end;
         end;
    except
        on e: exception do
        AddLog('Error  : ' + e.Message + ' ** Trace: ' + e.StackTrace);
    end;
end;

          SincronizarSAP(); // Queda pendiente hasta nueva solicitud
          ActualizarMySQL();
          ActualizarSAP();


end.


