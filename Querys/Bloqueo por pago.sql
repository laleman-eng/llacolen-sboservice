Select DISTINCT spr.cardcode SocPrin, s.CardCode, s.LicTradNum, s.GroupCode, spr.saldo, s.frozenFor, s.U_Tickets
  from OCRD s, 
       (select i.CardCode, i.DocTotal-i.PaidToDate as Saldo, i.DocDueDate, 
               c.U_cod , c.U_cod1, c.U_co2 , c.U_cod3, c.U_cod4 , c.U_cod5, 
               c.U_cod6, c.U_cod7, c.U_cod8, c.U_cod9, c.U_cod10, c.U_cod11 
          from OINV i inner join OCRD c  on i.CardCode = c.CardCode 
         where i.DocTotal > i.PaidToDate  
           and i.DocDueDate < GETDATE()   
           and c.GroupCode = 100 
           and c.CardType  = 'C') spr 
 where s.CardCode = spr.CardCode   
    or s.CardCode = spr.U_cod   
    or s.CardCode = spr.U_cod1  
    or s.CardCode = spr.U_co2   
    or s.CardCode = spr.U_cod3  
    or s.CardCode = spr.U_cod4  
    or s.CardCode = spr.U_cod5  
    or s.CardCode = spr.U_cod6  
    or s.CardCode = spr.U_cod7  
    or s.CardCode = spr.U_cod8  
    or s.CardCode = spr.U_cod9  
    or s.CardCode = spr.U_cod10 
    or s.CardCode = spr.U_cod11 
 order by spr.CardCode
