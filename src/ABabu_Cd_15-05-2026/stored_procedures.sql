
-- ============================================================
-- PROCEDURE: PrcPrepareAccLedger
-- ============================================================
CREATE  procedure [dbo].[PrcPrepareAccLedger] (
    @VFirm nchar(6),
    @From_dt date,
    @To_dt date,
    @Accode nvarchar(max)=null,   
    @Agcode nvarchar(max)=null,  
    @Ucode nchar(20),
	@IsDepotFirm int,
	@CGSTAc nvarchar(max)=null,
	

-- ============================================================
-- PROCEDURE: PrcPrepareAccLedgerBrokSale
-- ============================================================
CREATE  procedure [dbo].[PrcPrepareAccLedgerBrokSale] (
    @VFirm nchar(6),
    @From_dt date,
    @To_dt date,
	@YrEnd_dt date,
    @Accode nvarchar(max)=null,   
    @Agcode nvarchar(max)=null,  
    @BrokkerageAc nvarchar(max)=null,  
	@BrkTp

-- ============================================================
-- PROCEDURE: PrcPrepareAccLedgerBrokSaleOpBal
-- ============================================================
CREATE  procedure [dbo].[PrcPrepareAccLedgerBrokSaleOpBal] (
    @VFirm nchar(6),
    @From_dt date,
    @To_dt date,
	@YrEnd_dt date,
    @Accode nvarchar(max)=null,   
    @Agcode nvarchar(max)=null,  
    @BrokkerageAc nvarchar(max)=null,  
	@

-- ============================================================
-- PROCEDURE: PrcPrepareBookingPartyVsDesp
-- ============================================================
create  procedure [dbo].[PrcPrepareBookingPartyVsDesp] (
    @VFirm nchar(6),
    @Fr_dt date,
    @To_dt date,
    @BkNo nvarchar(max)=null,   
    @BkYear nvarchar(max)=null,   
    @BkSrNo nvarchar(max)=null,   
    @Accode nvarchar(max)=null, 

-- ============================================================
-- PROCEDURE: PrcPrepareBrokSale
-- ============================================================
CREATE procedure [dbo].[PrcPrepareBrokSale] (
    @VFirm nchar(6),
	@Start_dt date,
    @From_dt date,
    @To_dt date,
    @Accode nvarchar(max)=null,   
    @Ucode nchar(20),
	@DepotFirm nchar(1),
	@BrkTp nchar(1),
	@VFirmAdd1 nchar(40),
	@VF

-- ============================================================
-- PROCEDURE: PrcPrepareClBalance
-- ============================================================
CREATE  procedure [dbo].[PrcPrepareClBalance] (
    @VFirm nchar(6),
    @Fr_dt date,
    @To_dt date,
	@FyStart_dt date,
    @Accode nvarchar(max)=null,   
    @Agcode nvarchar(max)=null,  
	@CGSTAc nvarchar(max)=null,
	@SGSTAc nvarchar(max)=n

-- ============================================================
-- PROCEDURE: PrcPrepareDailyEntryMillBill
-- ============================================================
CREATE procedure [dbo].[PrcPrepareDailyEntryMillBill] (
    @VFirm nchar(6),
    @Fr_dt date,
    @To_dt date,
    @Accode nvarchar(max)=null,   
    @Accode_Br nvarchar(max)=null,
	@MillCode nvarchar(max)=null, 
	@ItemCode nvarchar(max)=null, 
	

-- ============================================================
-- PROCEDURE: PrcPrepareDailyInward
-- ============================================================
create  procedure [dbo].[PrcPrepareDailyInward] (
    @VFirm nchar(6),
    @Fr_dt date,
    @To_dt date,
	@MillCode nvarchar(max)=null, 
	@ItemCode nvarchar(max)=null, 
    @Ucode nchar(20)
	
)
as
    set nocount on;
    declare @i1 int
   

-- ============================================================
-- PROCEDURE: PrcPrepareDailyReport
-- ============================================================
CREATE  procedure [dbo].[PrcPrepareDailyReport] (
    @Fr_dt date,
    @To_dt date,
    @Ucode nchar(20),
	@CYear nvarchar(max)=null,
	@FSr nvarchar(max)=null,
	@TSr nvarchar(max)=null,
	@PurBookFrom date
)
as
    set nocount on;
    declare @

-- ============================================================
-- PROCEDURE: PrcPrepareDailyReportMail
-- ============================================================
CREATE  procedure [dbo].[PrcPrepareDailyReportMail] (
	@YearStart_Dt as date,
    @Fr_dt date,
    @To_dt date,
	@BkFr_dt date,
	@MillCode nvarchar(max)=null, 
	@PtyGrpCode nvarchar(max)=null, 
    @Ucode nchar(20)
	
)
as
    set nocount on;

-- ============================================================
-- PROCEDURE: PrcPrepareDespatchDetail
-- ============================================================
CREATE  procedure [dbo].[PrcPrepareDespatchDetail] (
    @VFirm nchar(6),
    @Fr_dt date,
    @To_dt date,
    @Accode nvarchar(max)=null,   
    @Accode_Br nvarchar(max)=null,
	@MillCode nvarchar(max)=null, 
	@ItemCode nvarchar(max)=null, 
	@Vt

-- ============================================================
-- PROCEDURE: PrcPrepareGetPassWiseSale
-- ============================================================
CREATE  procedure [dbo].[PrcPrepareGetPassWiseSale] (
    @VFirm nchar(6),
    @Fr_dt date,
    @To_dt date,
    @Accode nvarchar(max)=null,   
    @Accode_Br nvarchar(max)=null,
	@MillCode nvarchar(max)=null, 
	@ItemCode nvarchar(max)=null,   
   

-- ============================================================
-- PROCEDURE: PrcPrepareGodownDelivery
-- ============================================================
CREATE   procedure [dbo].[PrcPrepareGodownDelivery] (
    @VFirm nchar(6),
    @Fr_dt date,
    @To_dt date,
    @GodCd nvarchar(max)=null,   
    @Accode nvarchar(max)=null,   
    @Accode_Br nvarchar(max)=null,
	@MillCode nvarchar(max)=null, 
	@I

-- ============================================================
-- PROCEDURE: PrcPrepareGodownDeliveryPending
-- ============================================================
CREATE   procedure [dbo].[PrcPrepareGodownDeliveryPending] (
    @VFirm nchar(6),
    @Fr_dt date,
    @To_dt date,
    @GodCd nvarchar(max)=null,   
    @Accode nvarchar(max)=null,   
    @Accode_Br nvarchar(max)=null,
	@MillCode nvarchar(max)=null

-- ============================================================
-- PROCEDURE: PrcPrepareGSTR2A
-- ============================================================
CREATE  procedure [dbo].[PrcPrepareGSTR2A] (
    @VFirm nchar(6),
    @Fr_dt date,
    @To_dt date,
    @Ucode nchar(20),
    @Accode nvarchar(max)=null   
)
as
    set nocount on;
    declare @i1 int
    BEGIN
        delete from TmpG

-- ============================================================
-- PROCEDURE: PrcPrepareGSTR3B
-- ============================================================
CREATE   procedure [dbo].[PrcPrepareGSTR3B] (
    @VFirm nchar(6),
    @Fr_dt date,
    @To_dt date,
 --   @Accode nvarchar(max)=null,   
 --   @Accode_Br nvarchar(max)=null,
	--@MillCode nvarchar(max)=null, 
	--@ItemCode nvarchar(max)=null,

-- ============================================================
-- PROCEDURE: PrcPrepareIntDebitNote
-- ============================================================
CREATE  procedure [dbo].[PrcPrepareIntDebitNote] (
    @VFirm nchar(6),
    @From_dt date,
    @To_dt date,
    @Accode nvarchar(max)=null,   
    @Accode_Br nvarchar(max)=null,  
    @Ucode nchar(20),
	@Grace nvarchar(max)=null,
	@IntTP nc

-- ============================================================
-- PROCEDURE: PrcPrepareLotwiseGodownStock
-- ============================================================
CREATE  procedure [dbo].[PrcPrepareLotwiseGodownStock] (
    @VFirm nchar(6),
    @Fr_dt date,
    @To_dt date,
	@MillCode nvarchar(max)=null, 
	@ItemCode nvarchar(max)=null, 
	@GodownCd nvarchar(max)=null, 
	@LotNo nchar(15), 
	@StkTP nchar(1), 

-- ============================================================
-- PROCEDURE: PrcPrepareLotwiseGodownStockDetail
-- ============================================================
CREATE procedure [dbo].[PrcPrepareLotwiseGodownStockDetail] (
    @VFirm nchar(6),
    @Fr_dt date,
    @To_dt date,
	@MillCode nvarchar(max)=null, 
	@ItemCode nvarchar(max)=null, 
	@GodownCd nvarchar(max)=null, 
	@LotNo nchar(15), 
	@StkTP nchar

-- ============================================================
-- PROCEDURE: PrcPrepareLotwiseStock
-- ============================================================
create  procedure [dbo].[PrcPrepareLotwiseStock] (
    @VFirm nchar(6),
    @Fr_dt date,
    @To_dt date,
	@MillCode nvarchar(max)=null, 
	@ItemCode nvarchar(max)=null, 
	@LotNo nchar(15), 
	@StkTP nchar(1), 
    @Vno nvarchar(max)=null,   
	@Vt

-- ============================================================
-- PROCEDURE: PrcPrepareMillStatement
-- ============================================================
CREATE  procedure [dbo].[PrcPrepareMillStatement] (
    @VFirm nchar(6),
    @Fr_dt date,
    @To_dt date,
	@YearStart_Dt as date,
	@MillCode nvarchar(max)=null, 
    @CommRt nvarchar(Max)=null,
	@CommOn nvarchar(Max)=null,
	@CashDiscRt nvarchar(

-- ============================================================
-- PROCEDURE: PrcPrepareMillStatementMillBill
-- ============================================================
CREATE  procedure [dbo].[PrcPrepareMillStatementMillBill] (
    @VFirm nchar(6),
    @Fr_dt date,
    @To_dt date,
	@MillCode nvarchar(max)=null, 
    @CommRt nvarchar(Max)=null,
	@CommOn nvarchar(Max)=null,
	@SrvTxRt nvarchar(Max)=null,
	@SWCess

-- ============================================================
-- PROCEDURE: PrcPrepareOpBalance
-- ============================================================
CREATE  procedure [dbo].[PrcPrepareOpBalance] (
    @VFirm nchar(6),
    @To_dt date,
	@FyStart_dt date,
    @Accode nvarchar(max)=null,   
    @Agcode nvarchar(max)=null,  
	@CGSTAc nvarchar(max)=null,
	@SGSTAc nvarchar(max)=null,
	@IGSTAc nvarc

-- ============================================================
-- PROCEDURE: PrcPrepareOpBalanceLedger
-- ============================================================
CREATE procedure [dbo].[PrcPrepareOpBalanceLedger] (
    @VFirm nchar(6),
    @To_dt date,
	@FyStart_dt date,
    @Accode nvarchar(max)=null,   
    @Agcode nvarchar(max)=null,  
	@CGSTAc nvarchar(max)=null,
	@SGSTAc nvarchar(max)=null,
	@IGSTA

-- ============================================================
-- PROCEDURE: PrcPrepareOutStangingSale
-- ============================================================
CREATE  procedure [dbo].[PrcPrepareOutStangingSale] (
    @VFirm nchar(6),
    @To_dt date,
    @Accode nvarchar(max)=null,   
    @Accode_Br nvarchar(max)=null,
	@IntDbAc nvarchar(max)=null,
	@IntDbAcWoGST nvarchar(max)=null,
	@IntTDSAc nvarchar(

-- ============================================================
-- PROCEDURE: PrcPrepareOutStangingSale_old
-- ============================================================
CREATE  procedure [dbo].[PrcPrepareOutStangingSale] (
    @VFirm nchar(6),
    @To_dt date,
    @Accode nvarchar(max)=null,   
    @Accode_Br nvarchar(max)=null,
	@IntDbAc nvarchar(max)=null,
	@IntTDSAc nvarchar(max)=null,
	@IntCrAc nvarchar(max)=

-- ============================================================
-- PROCEDURE: PrcPrepareOutStangingSaleDirectPayment
-- ============================================================
create procedure [dbo].[PrcPrepareOutStangingSaleDirectPayment] (
    @VFirm nchar(6),
    @To_dt date,
    @Accode nvarchar(max)=null,   
    @Accode_Br nvarchar(max)=null,
	@IntDbAc nvarchar(max)=null,
	@IntDbAcWoGST nvarchar(max)=null,
	@IntTDS

-- ============================================================
-- PROCEDURE: PrcPreparePendingBookingMillBill
-- ============================================================
create  procedure [dbo].[PrcPreparePendingBookingMillBill] (
    @VFirm nchar(6),
    @Fr_dt date,
    @To_dt date,
    @BkNo nvarchar(max)=null,   
    @BkYear nvarchar(max)=null,   
    @BkSrNo nvarchar(max)=null,   
    @Accode nvarchar(max)=nu

-- ============================================================
-- PROCEDURE: PrcPreparePendingBookingParty
-- ============================================================
CREATE  procedure [dbo].[PrcPreparePendingBookingParty] (
    @VFirm nchar(6),
    @Fr_dt date,
    @To_dt date,
    @BkNo nvarchar(max)=null,   
    @BkYear nvarchar(max)=null,   
    @BkSrNo nvarchar(max)=null,   
    @Accode nvarchar(max)=null,

-- ============================================================
-- PROCEDURE: PrcPreparePendingBookingPurch
-- ============================================================
CREATE  procedure [dbo].[PrcPreparePendingBookingPurch] (
    @VFirm nchar(6),
    @Fr_dt date,
    @To_dt date,
    @BkNo nvarchar(max)=null,   
    @BkYear nvarchar(max)=null,   
    @BkSrNo nvarchar(max)=null,   
    @Accode nvarchar(max)=null,

-- ============================================================
-- PROCEDURE: PrcPreparePurchaseDetail
-- ============================================================
CREATE  procedure [dbo].[PrcPreparePurchaseDetail] (
    @VFirm nchar(6),
    @Fr_dt date,
    @To_dt date,
    @Accode nvarchar(max)=null,   
    @Accode_Br nvarchar(max)=null,
	@MillCode nvarchar(max)=null, 
	@ItemCode nvarchar(max)=null, 
	@Vtyp

-- ============================================================
-- PROCEDURE: PrcPreparePurchaseDetailSIT
-- ============================================================
CREATE  procedure [dbo].[PrcPreparePurchaseDetailSIT] (
    @VFirm nchar(6),
    @Fr_dt date,
    @To_dt date,
    @Accode nvarchar(max)=null,   
    @Accode_Br nvarchar(max)=null,
	@MillCode nvarchar(max)=null, 
	@ItemCode

-- ============================================================
-- PROCEDURE: PrcPreparePurchaseGST
-- ============================================================
CREATE  procedure [dbo].[PrcPreparePurchaseGST] (
    @VFirm nchar(6),
    @Fr_dt date,
    @To_dt date,
    @Accode nvarchar(max)=null,   
    @Accode_Br nvarchar(max)=null,
	@MillCode nvarchar(max)=null, 
	@ItemCode nvarchar(max)=null, 
	@Vty

-- ============================================================
-- PROCEDURE: PrcPrepareRegisterCrDrNoteGST
-- ============================================================
CREATE  procedure [dbo].[PrcPrepareRegisterCrDrNoteGST] (
    @VFirm nchar(6),
    @From_dt date,
    @To_dt date,
    @Ucode nchar(20),
	@RegTp nchar(2)
)
as
    set nocount on;
    declare @i1 int
    BEGIN
        delete from TmpGentbl  whe

-- ============================================================
-- PROCEDURE: PrcPrepareRegisterJV
-- ============================================================
CREATE  procedure [dbo].[PrcPrepareRegisterJV] (
    @VFirm nchar(6),
    @From_dt date,
    @To_dt date,
    @Ucode nchar(20),
	@RegTp nchar(2)
)
as
    set nocount on;
    declare @i1 int
    BEGIN
        delete from TmpAccLedger  where Use

-- ============================================================
-- PROCEDURE: PrcPrepareRegisterLpInt
-- ============================================================
CREATE  procedure [dbo].[PrcPrepareRegisterLpInt] (
    @VFirm nchar(6),
    @From_dt date,
    @To_dt date,
    @Ucode nchar(20),
	@RegTp nchar(2)
)
as
    set nocount on;
    declare @i1 int
    BEGIN
        delete from TmpGentbl  where Use

-- ============================================================
-- PROCEDURE: PrcPrepareSaleDetailSIT
-- ============================================================
CREATE  procedure [dbo].[PrcPrepareSaleDetailSIT] (
    @VFirm nchar(6),
    @Fr_dt date,
    @To_dt date,
    @Accode nvarchar(max)=null,   
    @Accode_Br nvarchar(max)=null,
	@MillCode nvarchar(max)=null, 
	@ItemCode nva

-- ============================================================
-- PROCEDURE: PrcPrepareSaleItemDetail
-- ============================================================
CREATE  procedure [dbo].[PrcPrepareSaleItemDetail] (
    @VFirm nchar(6),
    @Fr_dt date,
    @To_dt date,
    @Accode nvarchar(max)=null,   
    @Accode_Br nvarchar(max)=null,
	@MillCode nvarchar(max)=null, 
	@ItemCode nvarchar(max)=null, 
    

-- ============================================================
-- PROCEDURE: PrcPrepareSaleRegister
-- ============================================================
CREATE   procedure [dbo].[PrcPrepareSaleRegister] (
    @VFirm nchar(6),
    @Fr_dt date,
    @To_dt date,
    @Accode nvarchar(max)=null,   
    @Accode_Br nvarchar(max)=null,
	@MillCode nvarchar(max)=null, 
	@ItemCode nvarchar(max)=null, 
	@V

-- ============================================================
-- PROCEDURE: PrcPrepareSaleRegisterGST
-- ============================================================
CREATE   procedure [dbo].[PrcPrepareSaleRegisterGST] (
    @VFirm nchar(6),
    @Fr_dt date,
    @To_dt date,
    @Accode nvarchar(max)=null,   
    @Accode_Br nvarchar(max)=null,
	@MillCode nvarchar(max)=null, 
	@ItemCode nvarchar(max)=null, 
	@

-- ============================================================
-- PROCEDURE: PrcPrepareSalesRegisterRCM
-- ============================================================
create  procedure [dbo].[PrcPrepareSalesRegisterRCM] (
    @VFirm nchar(6),
    @Fr_dt date,
    @To_dt date,
    @Accode nvarchar(max)=null,   
    @Accode_Br nvarchar(max)=null,
	@MillCode nvarchar(max)=null, 
	@ItemCode nvarchar(max)=null, 

-- ============================================================
-- PROCEDURE: PrcPrepareSaleStatus
-- ============================================================
CREATE  procedure [dbo].[PrcPrepareSaleStatus] (
    @VFirm nchar(6),
    @Fr_dt date,
    @To_dt date,
    @Accode nvarchar(max)=null,   
    @Accode_Br nvarchar(max)=null,
	@MillCode nvarchar(max)=null, 
    @Ucode nchar(20)
)
as
    set nocoun

-- ============================================================
-- PROCEDURE: PrcPrepareSaleStatusRatewise
-- ============================================================
CREATE  procedure [dbo].[PrcPrepareSaleStatusRatewise] (
    @VFirm nchar(6),
    @Fr_dt date,
    @To_dt date,
    @Accode nvarchar(max)=null,   
    @Accode_Br nvarchar(max)=null,
	@MillCode nvarchar(max)=null, 
	@ItemCode nvarchar(max)=null,   

-- ============================================================
-- PROCEDURE: PrcPrepareStock
-- ============================================================
CREATE  procedure [dbo].[PrcPrepareStock] (
    @VFirm nchar(6),
	@YearStart_Dt as date,
    @Fr_dt date,
    @To_dt date,
	@MillCode nvarchar(max)=null, 
	@ItemCode nvarchar(max)=null, 
    @Ucode nchar(20)
	
)
as
    set nocount on;
    dec

-- ============================================================
-- PROCEDURE: PrcPrepareStockDaily
-- ============================================================
CREATE  procedure [dbo].[PrcPrepareStockDaily] (
    @VFirm nchar(6),
	@YearStart_Dt as date,
    @Fr_dt date,
    @To_dt date,
	@MillCode nvarchar(max)=null, 
	@ItemCode nvarchar(max)=null, 
    @Ucode nchar(20)
	
)
as
    set nocount on;

-- ============================================================
-- PROCEDURE: PrcPrepareStockLedger
-- ============================================================
CREATE  procedure [dbo].[PrcPrepareStockLedger] (
    @VFirm nchar(6),
	@YearStart_Dt as date,
    @Fr_dt date,
    @To_dt date,
	@MillCode nvarchar(max)=null, 
	@ItemCode nvarchar(max)=null, 
    @Ucode nchar(20)
	
)
as
    set nocount on;
 

-- ============================================================
-- PROCEDURE: PrcPrepareSubAcShedule
-- ============================================================
CREATE  procedure [dbo].[PrcPrepareSubAcShedule] (
    @VFirm nchar(6),
    @From_dt date,
    @To_dt date,
	@AcCode nvarchar(max)=null,
	@ForAudit nvarchar(max)=null,
    @Ucode nchar(20)
)
as
    set nocount on;
    declare @i1 int
    BEGIN

-- ============================================================
-- PROCEDURE: PrcPrepareTcsReceivable
-- ============================================================
Create  procedure [dbo].[PrcPrepareTcsReceivable] (
    @VFirm nchar(6),
    @From_dt date,
    @To_dt date,
    @Accode nvarchar(max)=null,   
    @Agcode nvarchar(max)=null,  
    @Ucode nchar(20),
	@IsDepotFirm int,
	@TCSRec nvarchar(max)=null

-- ============================================================
-- PROCEDURE: PrcPrepareVatComputation
-- ============================================================
CREATE  procedure [dbo].[PrcPrepareVatComputation] (
    @VFirm nchar(6),
    @Fr_dt date,
    @To_dt date,
    @Ucode nchar(20)
	
)
as
    set nocount on;
    declare @i1 int
    BEGIN
        delete from TmpGentbl  where  UserName = @Ucode  

-- ============================================================
-- PROCEDURE: PrcPrepareVatTWSalePurch
-- ============================================================
CREATE  procedure [dbo].[PrcPrepareVatTWSalePurch] (
    @VFirm nchar(6),
    @Fr_dt date,
    @To_dt date,
	@RpTp  nchar(1),
    @Ucode nchar(20)
	
)
as
    set nocount on;
    declare @i1 int
    BEGIN
        delete from TmpGentbl  where  

-- ============================================================
-- PROCEDURE: PrcPrepareYearEnd
-- ============================================================
create  procedure [dbo].[PrcPrepareYearEnd] (
    @VFirm nchar(6),
    @To_dt date,
    @Accode nvarchar(max)=null,   
    @Accode_Br nvarchar(max)=null,
	@IntDbAc nvarchar(max)=null,
	@IntDbAcWoGST nvarchar(max)=null,
	@IntTDSAc nvarchar(max)=nul

