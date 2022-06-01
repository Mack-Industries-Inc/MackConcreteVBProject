Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Transactions
Imports System.Data
Imports System.Runtime.InteropServices
Imports Mongoose.Core.Common
Imports Mongoose.Core.DataAccess
Imports Mongoose.IDO
Imports Mongoose.IDO.Protocol
Imports Mongoose.IDO.DataAccess

Namespace ue_Mack_ProjCommCalc
    ' <IDOExtensionClass( "ue_Mack_ProjCommCalc" )>
    Partial Public Class ue_Mack_ProjCommCalc : Inherits IDOExtensionClass
        ' Add methods and event handlers here
        Public Function ue_Mack_ProjCommCalcSp(ByVal pCustNum As String, ByVal pInvNum As String, ByVal pInvSeq As String, ByVal pSite As String, ByRef pInfobar As String) As Integer

            'Dim results As DataTable = New DataTable("Mack_AntiDeferred")
            Dim appDB As ApplicationDB = Me.CreateApplicationDB()
            Dim cmd As IDbCommand = appDB.Connection.CreateCommand()
            Dim parm_Site, parm_InvNum, parm_InvSeq, parm_CustNum, parm_Infobar As IDbDataParameter

            cmd.CommandType = CommandType.Text
            cmd.CommandText = "

EXEC SetSiteSp @pSite, NULL 
Declare @Infobar    InfobarType

Declare 
@Severity  int 
, @InvHdrInvNum              InvNumType
, @InvHdrInvSeq              InvSeqType
, @InvHdrCustNum             CustNumType
, @InvHdrCustSeq             CustSeqType
, @InvHdrCoNum               CoNumType
, @InvHdrInvDate             DateType
, @InvHdrRowPointer    RowpointerType 
, @InvHdrCost                AmountType
, @InvHdrTotCommDue          AmountType
, @InvHdrCommCalc            AmountType
, @InvHdrCommBase            AmountType
, @InvHdrCommDue             AmountType
, @TLoopCounter  GenericNoType
, @TCurSlsman    SlsmanType
, @SlsmanRowPointer          RowPointerType
, @SlsmanSlsman              SlsmanType
, @SlsmanSlsmangr            SlsmanType
, @SlsmanRefNum              EmpVendNumType
, @TCommCalc     GenericDecimalType
, @TCommBase     GenericDecimalType
, @TCommBaseTot  GenericDecimalType
, @InvCred                 NCHAR(1)
, @CoCurrCode                CurrCodeType
, @CurrencyRowPointer        RowPointerType
, @CurrencyCurrCode          CurrCodeType
, @CurrencyPlaces            DecimalPlacesType
, @CurrencyPlacesCp          DecimalPlacesType
, @CoSlsCommSlsman           SlsmanType
, @CoSlsCommRevPercent       CommPercentType
, @CoSlsCommCommPercent      CommPercentType
, @CoSlsCommCoLine           CoLineType
, @InvHdrExchRate            ExchRateType
, @ParmsSite                 SiteType
, @InvHdrSlsman     SlsmanType
, @CommdueRowPointer         RowPointerType
, @CommdueInvNum             InvNumType
, @CommdueCoNum              CoNumType
, @CommdueSlsman             SlsmanType
, @CommdueCustNum            CustNumType
, @CommdueCommDue            AmountType
, @CommdueDueDate            DateType
, @CommdueCommCalc           AmountType
, @CommdueCommBase           AmountType
, @CommdueCommBaseSlsp       AmountType
, @CommdueSeq                CommdueSeqType
, @CommduePaidFlag           ListYesNoType
, @CommdueSlsmangr           SlsmanType
, @CommdueStat               CommdueStatusType
, @CommdueRef                ReferenceType
, @CommdueEmpNum             EmpNumType
, @CoparmsDueOnPmt           AmountType
, @TCrMemo       LongListType
, @TInvLabel     LongListType
, @InvHdrPrice    AmountType
, @InvHdrDisc          GenericDecimalType -- OrderDiscType
, @InvHdrDiscAmount AmountType
, @InvHdrFreight AmountType
, @InvHdrMiscCharges AmountType
, @InvHdrPrepaidAmt AmountType
, @SalesTax AmtTotType
, @BasePrice  AmountType
, @CommtabSalesOnlyRowPointer RowPointerType
, @CommtabSalesOnlyCvalue##1 CommtabValueType
, @CommtabWildcardRowPointer RowPointerType
, @CommtabWildcardCvalue##1 CommtabValueType
, @CommtabRowPointer   RowPointerType
, @CommtabCvalue##1    CommtabValueType
, @TRevPercent   CommPercentType
, @TCommPercent  CommPercentType

Set @TRevPercent = 100
Set @TCommPercent = NULL

DECLARE @TmpCommdue TABLE (
     inv_num                 InvNumType
   , co_num                  CoNumType
   , slsman                  SlsmanType
   , cust_num                CustNumType
   , comm_due                AmountType
   , due_date                DateType
   , comm_calc               AmountType
   , comm_base               AmountType
   , comm_base_slsp          AmountType
   , seq                     CommdueSeqType
   , paid_flag               ListYesNoType
   , slsmangr                SlsmanType
   , stat                    CommdueStatusType
   , ref                     ReferenceType
   , emp_num                 EmpNumType
   )

   DECLARE @TmpSlsman TABLE (
     sales_ytd               AmountType
   , sales_ptd               AmountType
   , RowPointer              RowPointerType PRIMARY KEY
   )

Set @Severity = 0
SET @TInvLabel = dbo.GetLabel('@inv_stax.inv_num')
SELECT @CoparmsDueOnPmt = coparms.due_on_pmt FROM coparms with (readuncommitted)

SELECT @ParmsSite = parms.site FROM parms with (readuncommitted)

SELECT
  @CommtabWildcardRowPointer = commtab.RowPointer
, @CommtabWildcardCvalue##1  = commtab.cvalue##1
FROM commtab with (readuncommitted)
WHERE commtab.field1 = '*'
AND commtab.field2 = '*'

SET @TLoopCounter  = 0

   DECLARE inv_hdr_Crs CURSOR LOCAL STATIC FOR
   SELECT
        proj_inv_hdr.inv_num
      , proj_inv_hdr.proj_num
      , inv_hdr.exch_rate
      , inv_hdr.curr_code
      , inv_hdr.RowPointer
      , inv_hdr.cust_num
      , inv_hdr.cust_seq
     -- , proj_inv_hdr.uf_slsman
   , inv_hdr.price
   , inv_hdr.inv_date
   ,proj_inv_hdr.type
   FROM proj_inv_hdr
   inner join inv_hdr on proj_inv_hdr.inv_num = inv_hdr.inv_num and proj_inv_hdr.inv_seq = inv_hdr.inv_seq
   where proj_inv_hdr.inv_num = @pInvNum and proj_inv_hdr.inv_seq = @pInvSeq and proj_inv_hdr.cust_num = @pCustNum
   OPEN inv_hdr_Crs
   WHILE @Severity = 0
   BEGIN
      FETCH inv_hdr_Crs INTO
           @InvHdrInvNum
         , @InvHdrCoNum
         , @InvHdrExchRate
         , @CurrencyCurrCode
         , @InvHdrRowPointer
         , @InvHdrCustNum
         , @InvHdrCustSeq
   --, @InvHdrSlsman
   , @InvHdrPrice
   , @InvHdrInvDate
   ,@InvCred
      IF @@FETCH_STATUS = -1
          BREAK

      SET @InvHdrCommBase = 0
      SET @InvHdrCommCalc = 0
      SET @InvHdrCommDue = 0
      SET @InvHdrTotCommDue = 0

   select @InvHdrSlsman = uf_slsman from proj where proj_num = @InvHdrCoNum
   select @CoCurrCode = curr_code from co where co_num = @InvHdrCoNum
   Set @CoCurrCode = 'USD'  --hw debug
   
      SELECT
        @CurrencyRowPointer = currency.RowPointer
      , @CurrencyCurrCode   = currency.curr_code
      , @CurrencyPlaces = currency.places
      , @CurrencyPlacesCp = currency.places_cp
      FROM currency with (readuncommitted)
      WHERE currency.curr_code = @CoCurrCode


/*
      DECLARE co_sls_comm_Crs CURSOR LOCAL STATIC FOR
      SELECT
          inv_item.co_line
      FROM inv_item
      WHERE inv_item.inv_num = @InvHdrInvNum

      OPEN co_sls_comm_Crs
      WHILE 0 = 0
      BEGIN
         FETCH co_sls_comm_Crs INTO
             @CoSlsCommCoLine
         IF @@FETCH_STATUS = -1
             BREAK
*/
         Set @CoSlsCommRevPercent = 100  
         Set @CoSlsCommCommPercent = 0 
         SET @CoSlsCommRevPercent  = ISNULL(@CoSlsCommRevPercent, 0)

         SET @TCurSlsman = @InvHdrSlsman
         SET @TLoopCounter = 0

         --While @TCurSlsman <> ''
         WHILE @TCurSlsman IS NOT NULL
         BEGIN
            SET @TLoopCounter = @TLoopCounter + 1
            IF @TLoopCounter > 20
            BEGIN
               SET @Infobar = NULL
               EXEC @Severity = dbo.MsgAppSp @Infobar OUTPUT, 'E=Recursive1'
                  , '@slsman'
                  , '@slsman.slsman'
                  , @TCurSlsman

               --GOTO EXIT_SP
            END

            SET @SlsmanRowPointer = NULL
            SET @SlsmanSlsman     = NULL
            SET @SlsmanSlsmangr   = NULL
            SET @SlsmanRefNum     = NULL

            SELECT
                 @SlsmanRowPointer = slsman.RowPointer
               , @SlsmanSlsman     = slsman.slsman
               , @SlsmanSlsmangr   = slsman.slsmangr
               , @SlsmanRefNum     = slsman.ref_num
            FROM slsman WITH (UPDLOCK)
            WHERE slsman.slsman = @TCurSlsman

            IF @SlsmanRowPointer IS NULL
            BEGIN
               SET @Infobar = NULL
               EXEC @Severity = dbo.MsgAppSp @Infobar OUTPUT, 'E=NoExist1'
                  , '@slsman'
                  , '@slsman.slsman'
                  , @TCurSlsman

               --hw GOTO EXIT_SP
            END

            SET @TCommCalc = 0
            SET @TCommBase = 0
            SET @TCommBaseTot = 0

   /*
            EXEC @Severity = dbo.CommCalcSp
                 @InvCred, /* Invoice/Credit Memo */
                 @InvHdrInvNum,                   /* Invoice Number */
                 @InvHdrCoNum,
                 @TCurSlsman,
                 @CurrencyCurrCode,
                 @CurrencyPlaces,
                 @InvHdrExchRate,
                 @CoSlsCommRevPercent,
                 @CoSlsCommCommPercent,
                 @CoSlsCommCoLine,
                 NULL,                            /* rowid for rma processing */
                 @TCommBaseTot OUTPUT,
                 @TCommCalc    OUTPUT,            /* Commission Amount */
                 @TCommBase    OUTPUT,            /* Commission Base */
                 @Infobar      OUTPUT
            , @ParmsSite = @ParmsSite

            IF @Severity <> 0
            BEGIN
               SET @Infobar = NULL
               EXEC @Severity = dbo.MsgAppSp @Infobar OUTPUT, 'E=CmdFailed1'
                  , '@commtran'
                  , '@co'
                  , '@co.co_num'
                  , @InvHdrCoNum
               --hw GOTO EXIT_SP
            END
   */

/*Cacluate commdue */   
  SELECT
    @InvHdrDisc = disc
  , @InvHdrPrice = price
  , @InvHdrDiscAmount = disc_amount
  , @InvHdrFreight = freight
  , @InvHdrMiscCharges = misc_charges
  , @InvHdrPrepaidAmt = prepaid_amt
  , @SalesTax = isnull((select sum(sales_tax) from inv_stax
     where inv_stax.inv_num = inv_hdr.inv_num
     and inv_stax.inv_seq = inv_hdr.inv_seq), 0)
  FROM inv_hdr
  WHERE inv_hdr.inv_num = @InvHdrInvNum AND inv_hdr.inv_seq = 0

  set @BasePrice = @InvHdrDiscAmount + @InvHdrPrice - @InvHdrFreight - @InvHdrMiscCharges - @SalesTax + @InvHdrPrepaidAmt

  SELECT
    @CommtabSalesOnlyRowPointer = commtab.RowPointer
  , @CommtabSalesOnlyCvalue##1  = commtab.cvalue##1
  FROM commtab with (readuncommitted)
  WHERE commtab.field1 = @InvHdrSlsman
  AND commtab.field2 = '*'

  SET @CommtabRowPointer = @CommtabSalesOnlyRowPointer
  SET @CommtabCvalue##1  = @CommtabSalesOnlyCvalue##1

  IF @CommtabRowPointer IS NULL
  BEGIN
   SET @CommtabRowPointer = @CommtabWildcardRowPointer
   SET @CommtabCvalue##1  = @CommtabWildcardCvalue##1
  END

  SET @TCommBaseTot = @BasePrice --@TCommBaseTot + @TBase
  --IF NOT @BaseTotOnly = 1
  BEGIN

   SET @TCommBase = @TCommBase + (@BasePrice * @TRevPercent / 100.0)
   IF @TCommPercent IS NULL
   SET @TCommCalc = (@CommtabCvalue##1 * @BasePrice / 100
       * @TRevPercent / 100.0)
   ELSE
   SET @TCommCalc = (@TCommPercent * @BasePrice / 100
       * @TRevPercent / 100.0)
  --hw debug select @BasePrice, @TCommPercent,@TCommCalc,@CommtabCvalue##1,@TRevPercent
  END
  
/*Cacluate commdue END*/

            SET @CoSlsCommCommPercent = NULL     -- Set commission percent to null for manager, which will cause the
                                                 -- manager to get paid according to the default % in their commtab
                                                 -- table record.

            SET @InvHdrTotCommDue = @InvHdrTotCommDue + @TCommCalc

            /* update inv-hdr with inv-hdr.slsman calculations only */
            IF ISNULL(@InvHdrSlsman, NCHAR(1)) = ISNULL(@TCurSlsman, NCHAR(1))
            BEGIN
               SET @InvHdrCommCalc = @InvHdrCommCalc + @TCommCalc
               SET @InvHdrCommBase = @TCommBase
               SET @InvHdrCommDue = @InvHdrCommDue +
                                    (CASE WHEN @CoparmsDueOnPmt=0 OR @InvCred = 'C'
                                          THEN @TCommCalc
                                          ELSE 0
                                          END)
            END

            IF @TCommCalc <> 0
            BEGIN
               --create commdue
               SET @CommdueInvNum       = @InvHdrInvNum
               SET @CommdueCoNum        = @InvHdrCoNum
               SET @CommdueSlsman       = @TCurSlsman
               SET @CommdueCustNum      = @InvHdrCustNum
               SET @CommdueCommDue      = (CASE WHEN @CoparmsDueOnPmt = 0 OR @InvCred = 'C'
                                                THEN @TCommCalc
                                                ELSE 0
                                                END)
               SET @CommdueDueDate      = @InvHdrInvDate
               SET @CommdueCommCalc     = @TCommCalc
               SET @CommdueCommBase     = @TCommBaseTot
               SET @CommdueCommBaseSlsp = @TCommBase
               SET @CommdueSeq          = 1
               SET @CommduePaidFlag     = 0
               SET @CommdueSlsmangr     = (CASE WHEN ISNULL(@SlsmanSlsman, NCHAR(1)) =
                                                     ISNULL(@SlsmanSlsmangr, NCHAR(1))
                                                     AND @TLoopCounter > 1
                                                THEN NULL
                                                ELSE @SlsmanSlsmangr
                                                END)
               SET @CommdueStat = 'P'
               SET @CommdueRef          = (CASE WHEN  @InvCred = 'I'
                                                  THEN @TInvLabel
                                                ELSE @TCrMemo
                                                END)
               SET @CommdueEmpNum       = @SlsmanRefNum

               INSERT INTO @TmpCommdue (
                    inv_num
                  , co_num
                  , slsman
                  , cust_num
                  , comm_due
                  , due_date
                  , comm_calc
                  , comm_base
                  , comm_base_slsp
                  , seq
                  , paid_flag
                  , slsmangr
                  , stat
                  , ref
                  , emp_num )
               VALUES(
                    @CommdueInvNum
                  , @CommdueCoNum
                  , @CommdueSlsman
                  , @CommdueCustNum
                  , @CommdueCommDue
                  , @CommdueDueDate
                  , @CommdueCommCalc
                  , @CommdueCommBase
                  , @CommdueCommBaseSlsp
                  , @CommdueSeq
                  , @CommduePaidFlag
                  , @CommdueSlsmangr
                  , @CommdueStat
                  , @CommdueRef
                  , @CommdueEmpNum )
            END /* IF @TCommCalc <> 0 */

            /* DON'T INCREASE TWICE
               IF HE MANAGES HIMSELF AND HE MADE THE SALE */

            IF @TLoopCounter <= 1 OR ISNULL(@CoSlsCommSlsman, NCHAR(1)) <> ISNULL(@SlsmanSlsmangr, NCHAR(1))
            BEGIN
               IF EXISTS (SELECT * FROM @TmpSlsman AS TS WHERE TS.RowPointer = @SlsmanRowPointer)
                  UPDATE @TmpSlsman
                     SET sales_ytd = sales_ytd + @TCommBase
                       , sales_ptd = sales_ptd + @TCommBase
                  WHERE RowPointer = @SlsmanRowPointer
               ELSE
                  INSERT INTO @TmpSlsman (
                       sales_ytd
                     , sales_ptd
                     , RowPointer)
                  SELECT
                       @TCommBase
                     , @TCommBase
                     , @SlsmanRowPointer
            END /* IF @TLoopCounter <= 1 ... */

            /* PROCESS MANAGER OF SALESMAN NEXT ITERATION, IF SALESMAN IS
               HIS/HER OWN MANAGER, ONLY PROCESS IF THE MANAGER MADE THE
               SALE THEMSELF (.ie t-loop-counter = 1). */
            SET @TCurSlsman =  CASE WHEN ISNULL(@SlsmanSlsman, NCHAR(1)) = ISNULL(@SlsmanSlsmangr, NCHAR(1)) AND
                                         @TLoopCounter > 1
                                    THEN NULL
                                    ELSE @SlsmanSlsmangr
                               END
         END /* WHILE @TCurSlsman IS NOT NULL */
      --END /* WHILE @Severity = 0 */
      --CLOSE      co_sls_comm_Crs
      --DEALLOCATE co_sls_comm_Crs /* for each co-sls-comm */

      UPDATE inv_hdr
         SET comm_base = @InvHdrCommBase
           , comm_calc = @InvHdrCommCalc
           , comm_due  = @InvHdrCommDue
           , tot_comm_due = @InvHdrTotCommDue
      WHERE RowPointer = @InvHdrRowPointer


   END
   CLOSE      inv_hdr_Crs
   DEALLOCATE inv_hdr_Crs


   select * from @TmpSlsman
   select * from @TmpCommdue


   INSERT INTO commdue (
        inv_num
      , co_num
      , slsman
      , cust_num
      , comm_due
      , due_date
      , comm_calc
      , comm_base
      , comm_base_slsp
      , seq
      , paid_flag
      , slsmangr
      , stat
      , ref
      , emp_num )
   SELECT
        inv_num
      , co_num
      , slsman
      , cust_num
      , comm_due
      , due_date
      , comm_calc
      , comm_base
      , comm_base_slsp
      , seq
      , paid_flag
      , slsmangr
      , stat
      , ref
      , emp_num
   FROM @TmpCommdue

           "

            parm_InvNum = cmd.CreateParameter()
            parm_InvNum.ParameterName = "@pInvNum"
            parm_InvNum.Value = pInvNum

            parm_InvSeq = cmd.CreateParameter()
            parm_InvSeq.ParameterName = "@pInvSeq"
            parm_InvSeq.Value = pInvSeq

            parm_CustNum = cmd.CreateParameter()
            parm_CustNum.ParameterName = "@pCustNum"
            parm_CustNum.Value = pCustNum

            parm_Site = cmd.CreateParameter()
            parm_Site.ParameterName = "@pSite"
            parm_Site.Value = pSite

            'parm_Infobar = cmd.CreateParameter()
            'parm_Infobar.ParameterName = "@Infobar"
            'parm_Infobar.Value = pInfobar


            cmd.Parameters.Add(parm_InvNum)
            cmd.Parameters.Add(parm_InvSeq)
            cmd.Parameters.Add(parm_CustNum)
            cmd.Parameters.Add(parm_Site)
            'cmd.Parameters.Add(parm_Infobar)

            cmd.Connection = appDB.Connection
            cmd.ExecuteNonQuery()
            'Dim reader As IDataReader = cmd.ExecuteReader()
            'results.Load(reader)

            'Return results.CreateDataReader()

        End Function

    End Class
End Namespace