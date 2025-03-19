--CREATE PROCEDURE [dbo].[spSendEmailODMotor]
--	@piIntEmail int
--AS
/****************************************************  
**  File             : spSendEmailODMotor.sql		**
**  Name             : spSendEmailODMotor           **
**  Desc             : MPX Report OD Motor          **
**  Author           : Stella						**
**  Date Time Create : 4 Mei 2017 18:00:05			**
****************************************************/ 	

BEGIN
	SET NOCOUNT ON;

    drop table if exists #ODPT
	drop table if exists #tempPT
	drop table if exists #ODBranch
	drop table if exists #tempRegional
	drop table if exists #wilayah
	drop table if exists #regional
	drop table if exists #odMAF
	drop table if exists #odMCF
	drop table if exists #ODBranchtemp
	drop table if exists #ODBranchlastmonth	
	drop table if exists #temp
    drop table if exists #area
	drop table if exists #ODJaringan
	drop table if exists #ODcabangpos
	drop table if exists #ODParentBranch
	

	if object_id('tempdb..#odMCF') is not null
		drop table #odMCF

	if object_id('tempdb..#odMAF') is not null
		drop table #odMAF

	if object_id('tempdb..#ODWilayahtemp') is not null
		drop table #ODWilayahtemp

	if object_id('tempdb..#ODPT') is not null
		drop table #ODPT
	
	if object_id('tempdb..#ODPTtemp') is not null
		drop table #ODPTtemp
	
	if object_id('tempdb..#ODRegional') is not null
		drop table #ODRegional
	
	if object_id('tempdb..#ODRegionaltemp') is not null
		drop table #ODRegionaltemp
	
	if object_id('tempdb..#ODWilayah') is not null
		drop table #ODWilayah

	if object_id('tempdb..#ODArea') is not null
		drop table #ODArea

	if object_id('tempdb..#ODAreaTemp') is not null
		drop table #ODAreaTemp

	if object_id('tempdb..#tempRegional') is not null
		drop table #tempRegional

	if object_id('tempdb..#ODBranch') is not null
		drop table #ODBranch

	if object_id('tempdb..#ODBranchTemp') is not null
		drop table #ODBranchTemp

    if object_id('tempdb..#ODParentBranchTemp') is not null 
	    drop table #ODParentBranchTemp 

    if object_id('tempdb..#tempcabangpos') is not null
		drop table #tempcabangpos

	if object_id('tempdb..#odbranchlastmonth') is not null
	drop table #odbranchlastmonth





    declare @CurrDate datetime,@idx int, @mtdThisMonth varchar(8), @mtdlastMonth varchar(8), @OpenNPPaktifMAF int, @OpenNPPODMAF int, @OpenPersenNPPODMAF float, 
	        @OpenNPPaktifMCF int, @OpenNPPODMCF int, @OpenPersenNPPODMCF float, @OpenNPPaktifNasional int, @OpenNPPODNasional int, @OpenPersenNPPODNasional float, @TotRecordMCF int,
			@TotODMCF int, @TotPersenODMCF float, @TotRecordMAF int, @TotODMAF int, @TotPersenODMAF float, @TotRecordNasional int, @TotODNasional int, @TotPersenODNasional float, @TotalOpenAktifReg int,
			@TotalOpenODReg int , @TotalOpenPersenODReg float,@TotalCloseAktifReg int, @TotalCloseODReg int , @TotalClosePersenODReg float, @TotalLastMonthOpenAktifReg int, @TotalLastMonthOpenODReg int,
			@TotalLastMonthOpenPersenODReg float,@TotalLastMonthCloseAktifReg int,@TotalLastMonthCloseODReg int, @TotalLastMonthClosePersenODReg float, @TotalLastMonthCloseAktifWilayah int,
			@TotalLastMonthCloseODWilayah int,@od int, @ParentBranchName varchar(30),@TotRecordNPP int,@TotODNPP int,@TotPersenODNPP float,@OpenNPPAktif int,@OpenNPPOD int,@OpenPersenNPPOD float,
			@CloseNPPAktif int,@CloseNPPOD int,@ClosePersenNPPOD float,@LastMonthOpenNPPAktif int,@LastMonthOpenNPPOD int,@LastMonthOpenPersenNPPOD float,@LastMonthCloseNPPAktif int,@LastMonthCloseNPPOD int,
			@LastMonthClosePersenNPPOD float,@Pemisah varchar(max),@MsgHdr1 varchar(max),@MsgBody1 varchar(max),@MsgHdr2 varchar(max),@MsgBody2 varchar(max),@MsgHdr3 varchar(max),@MsgBody3 varchar(max), @MsgHdr4 varchar(max),
            @MsgBody4 varchar(max),@MsgHdr5 varchar(max),@MsgBody5 varchar(max),@MsgHdr6 varchar(max),@MsgBody6 varchar(max),@MsgFtr varchar(max),@MailBody varchar(max),@CurrRec int,@TotRec int,@ProfileName varchar(20), 
			@BodyFormat varchar(20), @Subject varchar(50),@mailRecipients varchar(max),@mailCopyRecipients varchar(max),@mailBlindCopyRecipients varchar(max),@Attachments varchar(max),
			@sql varchar(max),@PT varchar(20),@OpenNPPAktifPT FLOAT, @OpenNPPODPT FLOAT, @OpenPersenNPPODPT FLOAT, @CloseNPPAktifPT FLOAT, @CloseNPPODPT  FLOAT, @ClosePersenNPPODPT  FLOAT, @LastMonthOpenNPPAktifPT FLOAT, @LastMonthOpenNPPODPT  FLOAT, @TotalOpenAktifPT INT, @TotalOpenODPT INT, 
		    @TotalCloseAktifPT INT, @TotalCloseODPT INT, @TotalLastMonthOpenAktifPT INT, @TotalLastMonthOpenODPT INT, @TotalLastMonthCloseAktifPT INT, @TotalLastMonthCloseODPT INT, @TotalOpenPersenODPT FLOAT , @TotalClosePersenODPT FLOAT, @TotalLastMonthOpenPersenODPT FLOAT, @TotalLastMonthClosePersenODPT FLOAT,
		    @LastMonthOpenPersenNPPODPT  FLOAT, @LastMonthCloseNPPAktifPT  FLOAT, @LastMonthCloseNPPODPT  FLOAT, @LastMonthClosePersenNPPODPT  FLOAT, 
		    @OpenNPPAktifwilayah  FLOAT, @OpenNPPODwilayah FLOAT, @OpenPersenNPPODwilayah FLOAT, @CloseNPPAktifwilayah FLOAT, @CloseNPPODwilayah  FLOAT, @ClosePersenNPPODwilayah  FLOAT, @LastMonthOpenNPPAktifwilayah  FLOAT, @TotalOpenAktifwilayah INT, @TotalOpenODwilayah INT, 
		    @HBAname varchar(30),@TotalCloseAktifwilayah INT, @TotalCloseODwilayah INT, @TotalLastMonthOpenAktifwilayah INT, @TotalLastMonthOpenODwilayah INT, @LastMonthOpenNPPODwilayah  FLOAT, @LastMonthOpenPersenNPPODwilayah  FLOAT, @LastMonthCloseNPPAktifwilayah  FLOAT, @LastMonthCloseNPPODwilayah  FLOAT, @LastMonthClosePersenNPPODwilayah  FLOAT,
		    @TotalOpenPersenODwilayah FLOAT, @TotalClosePersenODwilayah FLOAT, @TotalLastMonthOpenPersenODwilayah FLOAT, @TotalLastMonthClosePersenODwilayah FLOAT,
		    @RegionalName VARCHAR(255), @wilayahname VARCHAR(255), @OpenNPPAktifregional  FLOAT, @OpenNPPODregional FLOAT, @OpenPersenNPPODregional FLOAT, @CloseNPPAktifregional FLOAT, @CloseNPPODregional  FLOAT, @ClosePersenNPPODregional  FLOAT, @LastMonthOpenNPPAktifregional  FLOAT, @LastMonthOpenNPPODregional  FLOAT,       
		    @LastMonthOpenPersenNPPODregional  FLOAT, @LastMonthCloseNPPAktifregional  FLOAT, @LastMonthCloseNPPODregional  FLOAT, @LastMonthClosePersenNPPODregional  FLOAT, @TotalOpenAktifRegional INT, @TotalOpenODRegional INT, @TotalCloseAktifRegional INT, @TotalCloseODRegional INT, @TotalLastMonthOpenAktifRegional INT, 
		    @TotalLastMonthOpenODRegional INT, @TotalLastMonthCloseAktifRegional INT, @TotalLastMonthCloseODRegional INT, @TotalOpenPersenODRegional FLOAT, @TotalClosePersenODRegional FLOAT, @TotalLastMonthOpenPersenODRegional FLOAT, @TotalLastMonthClosePersenODRegional FLOAT, 
		    @areaname VARCHAR(255),@OpenNPPAktifarea  FLOAT, @OpenNPPODarea FLOAT, @OpenPersenNPPODarea FLOAT, @CloseNPPAktifarea FLOAT, @CloseNPPODarea  FLOAT, @ClosePersenNPPODarea  FLOAT, @LastMonthOpenNPPAktifarea  FLOAT, @LastMonthOpenNPPODarea  FLOAT, @LastMonthOpenPersenNPPODarea  FLOAT, @TotalOpenAktifarea INT, @TotalOpenODarea INT,
	        @LastMonthCloseNPPAktifarea  FLOAT, @LastMonthCloseNPPODarea  FLOAT, @LastMonthClosePersenNPPODarea  FLOAT, @TotalCloseAktifarea FLOAT, @TotalCloseODarea FLOAT, @TotalLastMonthOpenAktifarea INT, @TotalLastMonthOpenODarea INT, @TotalLastMonthCloseAktifarea INT , @TotalLastMonthCloseODarea INT , @TotalOpenPersenODarea FLOAT, @TotalClosePersenODarea FLOAT, 
		    @TotalLastMonthOpenPersenODarea FLOAT, @TotalLastMonthClosePersenODarea FLOAT, @parentbranch VARCHAR(255), @TotalOpenAktifparentbranch INT, @TotalOpenODparentbranch INT, @TotalCloseAktifparentbranch INT, @TotalCloseODparentbranch INT, @TotalLastMonthOpenAktifparentbranch INT, @TotalLastMonthOpenODparentbranch INT, @TotalLastMonthCloseAktifparentbranch INT, @TotalLastMonthCloseODparentbranch INT,
		    @TotalClosePersenODparentbranch FLOAT, @TotalLastMonthOpenPersenODparentbranch FLOAT, @TotalLastMonthClosePersenODparentbranch FLOAT, @TotalOpenPersenODparentbranch FLOAT,
		    @OpenNPPAktifparentbranch  FLOAT, @OpenNPPODparentbranch FLOAT, @OpenPersenNPPODparentbranch FLOAT, @CloseNPPAktifparentbranch FLOAT, @CloseNPPODparentbranch  FLOAT, @ClosePersenNPPODparentbranch  FLOAT, @LastMonthOpenNPPAktifparentbranch  FLOAT,    
		    @LastMonthOpenNPPODparentbranch  FLOAT, @LastMonthOpenPersenNPPODparentbranch  FLOAT, @LastMonthCloseNPPAktifparentbranch  FLOAT, @LastMonthCloseNPPODparentbranch  FLOAT, @LastMonthClosePersenNPPODparentbranch  FLOAT,		
		    @branchname VARCHAR(255), @OpenNPPAktifbranchname  FLOAT, @OpenNPPODbranchname FLOAT, @OpenPersenNPPODbranchname FLOAT, @CloseNPPAktifbranchname FLOAT, @CloseNPPODbranchname  FLOAT, @ClosePersenNPPODbranchname  FLOAT,         
		    @LastMonthOpenNPPAktifbranchname  FLOAT, @LastMonthOpenNPPODbranchname  FLOAT, @LastMonthOpenPersenNPPODbranchname  FLOAT, @LastMonthCloseNPPAktifbranchname  FLOAT, @LastMonthCloseNPPODbranchname  FLOAT, @LastMonthClosePersenNPPODbranchname  FLOAT, @TotalOpenAktifBranchName INT, @TotalOpenODBranchName INT ,
		    @TotalCloseAktifBranchName INT, @TotalCloseODBranchName INT, @TotalLastMonthOpenAktifBranchName INT, @TotalLastMonthOpenODBranchName INT, @TotalLastMonthCloseAktifBranchName INT, @TotalLastMonthCloseODBranchName INT, @TotalClosePersenODBranchName FLOAT, @TotalLastMonthOpenPersenODBranchName FLOAT, @TotalLastMonthClosePersenODBranchName FLOAT, @TotalOpenPersenODBranchName FLOAT, 
			@branchid varchar (225), @piIntEmail int = '2' ;

set @branchid ='456'



create table #ODBranch(
			BranchID varchar(100),
			BranchName varchar(100),
			OpenNPPaktif numeric(18,0),
			OpenNPPod numeric(18,0),
			OpenPersenNPPOD float,
			CloseNPPaktif numeric(18,0),
			CloseNPPod numeric(18,0),
			ClosePersenNPPOD float,
			LastMonthOpenNPPaktif numeric(18,0),
			LastMonthOpenNPPod numeric(18,0),
			LastMonthOpenPersenNPPOD float,
			LastMonthCloseNPPaktif numeric(18,0),
			LastMonthCloseNPPod numeric(18,0),
			LastMonthClosePersenNPPOD float)


  create table #odMAF (	
			WilayahID varchar(2),
			WilayahName varchar(25),
			RegionalID varchar(2),
			AreaID varchar(2),
			AreaName varchar(30),
			BranchID varchar(5),
			nppno varchar(10),
			approvedate datetime,
			OP numeric(21,2),
			DeadLine datetime,
			LastPaydate datetime,
			od int)


  create table #odMCF (	
			WilayahID varchar(2),
			WilayahName varchar(25),
			RegionalID varchar(2),
			AreaID varchar(2),
			AreaName varchar(30),
			BranchID varchar(5),
			nppno varchar(10),
			approvedate datetime,
			OP numeric(21,2),
			DeadLine datetime,
			LastPaydate datetime,
			od int)

  create table #ODPT(
			PT varchar(10),
			OpenNPPaktif numeric(18,0),
			OpenNPPod numeric(18,0),
			OpenPersenNPPOD float,
			CloseNPPaktif numeric(18,0),
			CloseNPPod numeric(18,0),
			ClosePersenNPPOD float,
			LastMonthOpenNPPaktif numeric(18,0),
			LastMonthOpenNPPod numeric(18,0),
			LastMonthOpenPersenNPPOD float,
			LastMonthCloseNPPaktif numeric(18,0),
			LastMonthCloseNPPod numeric(18,0),
			LastMonthClosePersenNPPOD float)

	--create table #tempcabangpos(
	--				 PT varchar(225),
	--				 HBA varchar(225),
	--				 HBA_email varchar(225),
	--				 wilayahId varchar(225),
	--				 wilayahname varchar(225),
	--				 wilayah_email varchar(225),
	--				 regionalid varchar(225),
	--				 regionalname varchar(225),
	--				 regional_email varchar(225),
	--				 areaid varchar(225),
	--				 arename varchar(225),
	--				 area_email varchar(225),
	--				 jaringanId varchar(225),
	--				 jaringaname varchar(225),
	--				 jaringan_email varchar(225),
	--				 branchid varchar(225),
	--				 branchname varchar(225),
	--				 branch_email varchar(225),
	--				 isactive int,
	--				 Opening_Date datetime,
	--				 Closing_Date datetime,
	--				 Last_Updated datetime)


END

	IF DATEPART(HOUR,GETDATE())='0'
		begin
			set @CurrDate=getdate()-1
		end
	else
		begin
			set @CurrDate=getdate()
		end


    set @mtdThisMonth = convert(varchar(8),@CurrDate,112)

	if @mtdThisMonth=convert(varchar(8),eomonth(@CurrDate),112) 
	   begin
	      select @mtdlastMonth = convert(varchar(8),eomonth(dateadd(m,-1,@mtdThisMonth)),112)
		    --select @mtdlastMonth = convert(varchar(8),eomonth(dateadd(m,-4,@mtdThisMonth)),112) --20240513 Req Risk 
	   end
    else
	   begin
	      select @mtdlastMonth = convert(varchar(8),dateadd(m,-1,@mtdThisMonth),112)
		    --select @mtdlastMonth = convert(varchar(8),dateadd(m,-4,@mtdThisMonth),112) --20240513 Req Risk 
	   end


	------------------------------------------Cabang Pos ----------------------------------------------------------------------------------
	----select * from #tempcabangpos
	----insert into  #tempcabangpos
	--select 'MAF', 'HBA1','anthony.wijaya@mcf.co.id','2','wilayah 4-6','anthony.wijaya@mcf.co.id','24','wilayah 6','anthony.wijaya@mcf.co.id',
	--		'98','jawa barat 3','anthony.wijaya@mcf.co.id','456','bandung 6','anthony.wijaya@mcf.co.id','456','bandung 6','anthony.wijaya@mcf.co.id',
	--		'1',getdate(),getdate(),getdate()
	--union all
	--select 'MAF', 'HBA1','anthony.wijaya@mcf.co.id','2','wilayah 4-6','anthony.wijaya@mcf.co.id','24','wilayah 6','anthony.wijaya@mcf.co.id',
	--		'98','jawa barat 3','anthony.wijaya@mcf.co.id','456','bandung 6','anthony.wijaya@mcf.co.id','106','pos bandung','anthony.wijaya@mcf.co.id',
	--		'1',getdate(),getdate(),getdate()


	--SELECT 
	--    c.PT, c.HBA,c.HBA_email,r.Wilayahid,r.Wilayahname,c.wilayah_email,r.regionalid,r.Regionalname,c.regional_email,
	--    a.areaid,a.areaname,c.area_email,c.jaringanId,c.jaringaname,c.jaringan_email,c.branchid,c.branchname, 
	--    c.branch_email,c.IsActive,c.Opening_Date,c.Closing_Date,c.Last_Updated
	---- INTO #tempRegional
	--FROM [MACF-DBSTG].[REPL_DBMCF_MACFDB].dbo.SysGFCompanyBranch b
	--     INNER JOIN [MACF-DBSTG].[REPL_DBMCF_MACFDB].dbo.SysGFCompanyArea a ON b.areaid = a.areaid AND b.CompanyID = a.CompanyID
	--     INNER JOIN [MACF-DBSTG].[REPL_DBMCF_MACFDB].dbo.SysGFCompanyRegional r ON a.regionalid = r.regionalid AND r.CompanyID = a.CompanyID
	--     INNER JOIN #temptcabangpos c ON b.branchid = c.branchid
	---- INNER JOIN [MACF-DBSTG].DUMP_MACF.dbo.SYSGFCompanyBranchR2_Oct22 br 
	---- ON br.BranchId = b.BranchId --add by Eka 20 Oct 2022
	--WHERE 
	--    a.IsActive = '1' 
	--    AND b.IsActive = '1'
	--    AND r.WilayahID IN ('1', '2', '6')
	--    AND r.RegionalID NOT IN ('1', '18', '27', '7', '28');  -- add by yohan (Regional id '28')





------------------------branch-------------------------------------------------------------------------------------------------	
	Select b.branchid, b.branchName,a.areaid,a.areaname,r.regionalid,r.wilayahid,r.wilayahname
		   Into #tempRegional
		   From [MACF-DBSTG].[REPL_DBMCF_MACFDB].dbo.SysGFCompanyBranch b
				Inner join [MACF-DBSTG].[REPL_DBMCF_MACFDB].dbo.SysGFCompanyArea a on b.areaid=a.areaid and b.CompanyID=a.CompanyID
				Inner join [MACF-DBSTG].[REPL_DBMCF_MACFDB].dbo.SysGFCompanyRegional r on a.regionalid=r.regionalid and r.CompanyID=a.CompanyID
				--Inner join [MACF-DBSTG].DUMP_MACF.dbo.SYSGFCompanyBranchR2_Oct22 br on br.BranchId=b.BranchId --add by Eka 20 Oct 2022
				WHERE a.IsActive='1' and b.IsActive='1'
					and r.WilayahID in ('1','2','6')
					and r.RegionalID not in ('1','18','27','7','28')

----------------------------------------  MCF ---------------------------------------------------------
	insert into #odMCF (WilayahID,WilayahName,RegionalID,AreaID,AreaName,BranchID,nppno,approvedate,OP) 
	select d.WilayahID,d.WilayahName,d.RegionalID,d.AreaID,d.AreaName, a.BranchID, a.nppno, a.Approvedate, b.op
		   from [MACF-DBSTG].[REPL_DBMCF_EPMCF].dbo.npp a with(nolock)
				left join [MACF-DBSTG].[REPL_DBMCF_EPMCF].dbo.cm b with(nolock) on a.nppno=b.NPPNo
				left Join [MACF-DBSTG].[REPL_DBMCF_EPMCF].dbo.InsClaim c with(nolock) on a.NPPNo = c.NPPNo  
				left join #tempRegional d with(nolock)on a.BranchId=d.BranchId
		 	    where b.status='1' and convert(char,a.approvedate,112)<=convert(char,@CurrDate,112) and a.StatusNpp='A' 
				      and a.ApproveDate is not null 
					  --and d.RegionalID is not null
					  and isnull(c.StatusClaim,'')=''
					  and isNull(a.IsInActive,0) = 0  



	update od 
		set DeadLine=CollectionDate
	from  #odMCF od
	inner join 
	(
		select nppno, CollectionDate=min(CollectionDate)
		from [MACF-DBSTG].[REPL_DBMCF_ECol].dbo.ec_arcard with(nolock)
		where convert(char,CollectionDate,112)<=convert(char,@CurrDate,112)
			and PayDate is null
		group by nppno
	) ar on od.nppno=ar.nppno

				 
	update od 
		set LastPaydate=@CurrDate
	from #odMCF od
	where DeadLine is not null

	update od 
		set od=datediff(day,isnull(DeadLine,@CurrDate),isnull(LastPaydate,@CurrDate))
	from #odMCF od

	select @TotRecordMCF=count(*) from #odMCF where RegionalID is not null
	select @TotODMCF=count(*) from #odMCF where od>0 and RegionalID is not null
	set @TotPersenODMCF =  round(CAST(@TotODMCF as float)/CAST(@TotRecordMCF as float)*100,2)


	---------------------------------  MAF -------------------------------------------------------
	insert into #odMAF (WilayahID,WilayahName,RegionalID,AreaID,AreaName,BranchID,nppno,approvedate,OP) 
	select d.WilayahID,d.WilayahName,d.RegionalID,d.AreaID,d.AreaName,a.BranchID, a.nppno, a.Approvedate, b.op
		   from [MACF-DBSTG].[REPL_DBMAF_EPMAF].dbo.npp a with(nolock)
				left join [MACF-DBSTG].[REPL_DBMAF_EPMAF].dbo.cm b with(nolock) on a.nppno=b.NPPNo
				left Join [MACF-DBSTG].[REPL_DBMAF_EPMAF].dbo.InsClaim c with(nolock) on a.NPPNo = c.NPPNo  
				left join #tempRegional d with(nolock)on a.BranchId=d.BranchId
		 	    where b.status='1' and convert(char,a.approvedate,112)<=convert(char,@CurrDate,112) and a.StatusNpp='A' 
				      and a.ApproveDate is not null 
					  and isnull(c.StatusClaim,'')='' 
					  and isNull(a.IsInActive,0) = 0  

	update od 
		set DeadLine=CollectionDate
	from  #odMAF od
	inner join 
	(
		select nppno, CollectionDate=min(CollectionDate)
		from [MACF-DBSTG].[REPL_DBMAF_ECol].dbo.ec_arcard with(nolock)
		where convert(char,CollectionDate,112)<=convert(char,@CurrDate,112)
			and PayDate is null
		group by nppno
	) ar on od.nppno=ar.nppno


	update od 
		set LastPaydate=@CurrDate
	from #odMAF od
	where DeadLine is not null

	update od 
		set od=datediff(day,isnull(DeadLine,@CurrDate),isnull(LastPaydate,@CurrDate))
	from #odMAF od

	select @TotRecordMAF=count(*) from #odMAF where RegionalID is not null
	select @TotODMAF=count(*) from #odMAF where od>0 and RegionalID is not null
	set @TotPersenODMAF =  round(CAST(@TotODMAF as float)/CAST(@TotRecordMAF as float)*100,2)



----------------OPEN----------------------------------------------
	set @sql = '
	insert into #ODBranch(BranchID, BranchName, OpenNPPaktif, OpenNPPod, OpenPersenNPPOD)
	select r.BranchID
		, r.BranchName
		, NPPAktif=count(od.NppNo)
		, NPPOD=sum(case when od.od>0 Then 1 else 0 end)
		, PersenNPPOD=round(cast(sum(case when od.od>0 Then 1 else 0 end) as float) / cast(count(od.NppNo) as float)*100,2)
	from [MACF-DBREP].MAILINFO.DBO.ODHistoryDetail od with(nolock) 
		inner join #tempRegional r with(nolock) on od.branchid=r.branchid
	where oddate=convert(varchar(8),'+@mtdThismonth+',112) and OpenClose=''1'' 
		and r.RegionalID not in (''8'',''0'',''1'',''18'',''27'',''7'',''28'')
	group by r.BranchID, r.BranchName


	
	' exec (@sql)
	

-----------------------------------------close --------------------------------------------------------------------------------------------------------------------


insert into #ODBranch (BranchID, BranchName,CloseNPPaktif, CloseNPPod, ClosePersenNPPOD)
	select BranchID, BranchName, CloseNPPaktif=sum(CloseNPPaktif), CloseNPPod=sum(CloseNPPod), ClosePersenNPPOD=sum(ClosePersenNPPOD)
	From (
		select BranchID, BranchName, OpenNPPaktif=0, OpenNPPOd=0, OpenPersenNPPOD=0, 
		CloseNPPaktif=sum(CloseNPPaktif), CloseNPPod=sum(CloseNPPod),
		ClosePersenNPPOD = round(CAST(sum(CloseNPPod) as float)/CAST(sum(CloseNPPaktif) as float)*100,2)	 
		from 
		(
			select d.BranchID, r.BranchName, CloseNPPaktif=count(*), CloseNPPod=0 
			from #odMAF d
			left join #tempRegional r on r.branchID = d.BranchID
			where d.RegionalID<>'8'
			group by d.BranchID, r.BranchName
		Union all
			select d.BranchID, r.BranchName, CloseNPPaktif=0, CloseNPPod=count(*) 
			from #odMAF d
			left join #tempRegional r on r.branchID = d.BranchID
			where d.od>0 
			and d.RegionalID<>'8'
			group by d.BranchID, r.BranchName
		Union all
			select d.BranchID, r.BranchName, CloseNPPaktif=count(*), CloseNPPod=0 
			from #odMCF d
			left join #tempRegional r on r.branchID = d.BranchID
			where d.RegionalID<>'8'
			group by d.BranchID, r.BranchName
		Union all
		select d.BranchID, r.BranchName, CloseNPPaktif=0, CloseNPPod=count(*) 
			from #odMCF d
			left join #tempRegional r on r.branchID = d.BranchID
			where d.od>0 
			and d.RegionalID<>'8'
			group by d.BranchID, r.BranchName
		) a
		group by BranchID, BranchName
	)Qry 
	group by BranchID, BranchName
	


print(@mtdlastMonth)


set @sql='
	insert into #ODBranch(BranchID, BranchName, LastMonthOpenNPPaktif, LastMonthOpenNPPod, LastMonthOpenPersenNPPOD)
	select r.BranchID
		, r.BranchName
		, NPPAktif=count(od.NppNo)
		, NPPOD=sum(case when od.od>0 Then 1 else 0 end)
		, PersenNPPOD=round(cast(sum(case when od.od>0 Then 1 else 0 end) as float) / cast(count(od.NppNo) as float)*100,2)
	from [MACF-DBREP].MAILINFO.DBO.ODHistoryDetail_'+left(@mtdlastMonth,6)+' od with(nolock) 
		inner join #tempRegional r with(nolock) on od.branchid=r.branchid
	where oddate=convert(varchar(8),'+@mtdlastMonth+',112) and OpenClose=''1'' 
		and r.RegionalID not in (''8'',''0'',''1'',''18'',''27'',''7'',''28'')--add by yohan (Regional id 28) 05/09/2024
	group by r.BranchID, r.BranchName
	
	insert into #ODBranch(BranchID, BranchName, LastMonthCloseNPPaktif, LastMonthCloseNPPod, LastMonthClosePersenNPPOD)
	select r.BranchID
		, r.BranchName
		, NPPAktif=count(od.NppNo)
		, NPPOD=sum(case when od.od>0 Then 1 else 0 end)
		, PersenNPPOD=round(cast(sum(case when od.od>0 Then 1 else 0 end) as float) / cast(count(od.NppNo) as float)*100,2)
	from [MACF-DBREP].MAILINFO.DBO.ODHistoryDetail_'+left(@mtdlastMonth,6)+' od with(nolock) 
		inner join #tempRegional r with(nolock) on od.branchid=r.branchid
	where oddate=convert(varchar(8),'+@mtdlastMonth+',112) and OpenClose=''0'' 
		and r.RegionalID not in (''8'',''0'',''1'',''18'',''27'',''7'',''28'')--add by yohan (Regional id 28) 05/09/2024
	group by r.BranchID, r.BranchName
	'
	exec (@sql)


select idx=identity(int,1,1), BranchID, BranchName, OpenNPPaktif=sum(OpenNPPaktif), OpenNPPod=sum(OpenNPPod), OpenPersenNPPOD=sum(OpenPersenNPPOD), 
	       CloseNPPaktif=sum(CloseNPPaktif), CloseNPPod=sum(CloseNPPod), ClosePersenNPPOD=sum(ClosePersenNPPOD),
		   LastMonthOpenNPPaktif=sum(LastMonthOpenNPPaktif), LastMonthOpenNPPod=sum(LastMonthOpenNPPod), 
		   LastMonthOpenPersenNPPOD=sum(LastMonthOpenPersenNPPOD), LastMonthCloseNPPaktif=sum(LastMonthCloseNPPaktif), 
		   LastMonthCloseNPPod=sum(LastMonthCloseNPPod), LastMonthClosePersenNPPOD=sum(LastMonthClosePersenNPPOD)  
	into #ODBranchtemp
	from #ODBranch 
	group by BranchID,BranchName
	order by cast(BranchID as int)



	select idx=identity(int,1,1), BranchID, BranchName, OpenNPPaktif=sum(OpenNPPaktif), OpenNPPod=sum(OpenNPPod), OpenPersenNPPOD=sum(OpenPersenNPPOD), 
	       CloseNPPaktif=sum(CloseNPPaktif), CloseNPPod=sum(CloseNPPod), ClosePersenNPPOD=sum(ClosePersenNPPOD),
		   LastMonthOpenNPPaktif=sum(LastMonthOpenNPPaktif), LastMonthOpenNPPod=sum(LastMonthOpenNPPod), 
		   LastMonthOpenPersenNPPOD=sum(LastMonthOpenPersenNPPOD), LastMonthCloseNPPaktif=sum(LastMonthCloseNPPaktif), 
		   LastMonthCloseNPPod=sum(LastMonthCloseNPPod), LastMonthClosePersenNPPOD=sum(LastMonthClosePersenNPPOD)  
	into #ODBranchlastmonth
	from #ODBranch 
	group by BranchID,BranchName
	order by cast(BranchID as int)

---------------------------------------------------------------------------------------------------------------------
SELECT 
    PTID = br.companyID, 
    PTName = IIF(br.companyID = '2', 'MAF', 'MCF') ,
    HBA = 'HBA ' + r.WilayahID,
    HBAName = r.WilayahName,
    RegionalID = a.regionalID,
    RegionalName = r.regionalName,
    AreaID = a.AreaID,
    AreaName = a.AreaName,
    JaringanID = br3.branchID,
    Jaringanname = CASE 
        WHEN br3.branchname LIKE '%Bandar Lampung%' OR br3.branchname LIKE '%Batam%' 
        THEN UPPER(br3.branchname)
        ELSE UPPER(REPLACE(REPLACE(br3.branchname, 'MCF', ''), 'MAF', ''))
    END ,
    c.BranchID,
    c.Branchname,
    OpenNPPaktif,
    OpenNPPod,
    OpenPersenNPPOD,
    CloseNPPaktif,
    CloseNPPod,
    ClosePersenNPPOD,
    LastMonthOpenNPPaktif,
    LastMonthOpenNPPod,
    LastMonthOpenPersenNPPOD,
    LastMonthCloseNPPaktif,
    LastMonthCloseNPPod,
    LastMonthClosePersenNPPOD
INTO #temp
select * FROM [macf-dbstg].dump_macf.dbo.tb_mpx_laporan_od_cabang_pos c
LEFT JOIN [macf-dbstg].repl_dbmcf_macfdb.dbo.sysgfcompanybranch br ON c.branchID = br.branchID
LEFT JOIN [macf-dbstg].repl_dbmcf_macfdb.dbo.sysgfcompanybranch br3 ON br.konsolBranchID = br3.branchid
LEFT JOIN [macf-dbstg].repl_dbmcf_macfdb.dbo.sysgfcompanyarea a ON a.areaID = br.areaID AND br.companyID = a.companyID
LEFT JOIN [MACF-DBSTG].[REPL_DBMCF_MACFDB].dbo.MsBranch mb ON mb.BranchID = c.BranchID
LEFT JOIN [macf-dbstg].repl_dbmcf_macfdb.dbo.sysgfcompanyregional r ON r.regionalID = a.regionalID AND r.companyID = a.companyID


----WHERE br3.branchName IN ('BANDUNG VI', 'Tuban');
----drop table  #temp
--select * from #temp
----SELECT * FROM #dk WHERE jaringanname IN ('BANDUNG VI', 'Tuban');




--------------------------PT-------------------------------------------------------------------------------

select  
        idx=identity(int,1,1),
        PT = PTName,
        ISNULL(SUM(OpenNPPaktif), 0) AS OpenNPPaktif,
        ISNULL(SUM(OpenNPPod), 0) AS OpenNPPod,
        ISNULL(ROUND(CAST(ISNULL(SUM(OpenNPPod), 0) AS FLOAT) / NULLIF(CAST(SUM(OpenNPPaktif) AS FLOAT), 0) * 100, 2), 0) AS OpenPersenNPPOD,
        ISNULL(SUM(CloseNPPaktif), 0) AS CloseNPPaktif,
        ISNULL(SUM(CloseNPPod), 0) AS CloseNPPod,
        ISNULL(ROUND(CAST(ISNULL(SUM(CloseNPPod), 0) AS FLOAT) / NULLIF(CAST(SUM(CloseNPPaktif) AS FLOAT), 0) * 100, 2), 0) AS ClosePersenNPPOD,
        ISNULL(SUM(LastMonthOpenNPPaktif), 0) AS LastMonthOpenNPPaktif,
        ISNULL(SUM(LastMonthOpenNPPod), 0) AS LastMonthOpenNPPod,
        ISNULL(ROUND(CAST(ISNULL(SUM(LastMonthOpenNPPod), 0) AS FLOAT) / NULLIF(CAST(SUM(LastMonthOpenNPPaktif) AS FLOAT), 0) * 100, 2), 0) AS LastMonthOpenPersenNPPOD,
        ISNULL(SUM(LastMonthCloseNPPaktif), 0) AS LastMonthCloseNPPaktif,
        ISNULL(SUM(LastMonthCloseNPPod), 0) AS LastMonthCloseNPPod,
        ISNULL(ROUND(CAST(ISNULL(SUM(LastMonthCloseNPPod), 0) AS FLOAT) / NULLIF(CAST(SUM(LastMonthCloseNPPaktif) AS FLOAT), 0) * 100, 2), 0) AS LastMonthClosePersenNPPOD
INTO #tempPT
FROM  #temp 

GROUP BY  PTName


--select * from #tempPT


--drop table #tempPT
----drop table #tempPT

-----------------------------wilayah----------------------------------------------------------------------------------
SELECT 
 
    HBAName ,
    ISNULL(SUM(OpenNPPaktif), 0) AS OpenNPPaktif,
    ISNULL(SUM(OpenNPPod), 0) AS OpenNPPod,
    ISNULL(ROUND(CAST(ISNULL(SUM(OpenNPPod), 0) AS FLOAT) / NULLIF(CAST(SUM(OpenNPPaktif) AS FLOAT), 0) * 100, 2), 0) AS OpenPersenNPPOD,
    ISNULL(SUM(CloseNPPaktif), 0) AS CloseNPPaktif,
    ISNULL(SUM(CloseNPPod), 0) AS CloseNPPod,
    ISNULL(ROUND(CAST(ISNULL(SUM(CloseNPPod), 0) AS FLOAT) / NULLIF(CAST(SUM(CloseNPPaktif) AS FLOAT), 0) * 100, 2), 0) AS ClosePersenNPPOD,
    ISNULL(SUM(LastMonthOpenNPPaktif), 0) AS LastMonthOpenNPPaktif,
    ISNULL(SUM(LastMonthOpenNPPod), 0) AS LastMonthOpenNPPod,
    ISNULL(ROUND(CAST(ISNULL(SUM(LastMonthOpenNPPod), 0) AS FLOAT) / NULLIF(CAST(SUM(LastMonthOpenNPPaktif) AS FLOAT), 0) * 100, 2), 0) AS LastMonthOpenPersenNPPOD,
    ISNULL(SUM(LastMonthCloseNPPaktif), 0) AS LastMonthCloseNPPaktif,
    ISNULL(SUM(LastMonthCloseNPPod), 0) AS LastMonthCloseNPPod,
    ISNULL(ROUND(CAST(ISNULL(SUM(LastMonthCloseNPPod), 0) AS FLOAT) / NULLIF(CAST(SUM(LastMonthCloseNPPaktif) AS FLOAT), 0) * 100, 2), 0) AS LastMonthClosePersenNPPOD
INTO #wilayah
FROM #temp
--where HBAName in (select HBAName from #temp where konsolbranchid = @konsolbranchid )
GROUP BY  
    HBAName




--select * from #wilayah
--drop table  #wilayah

-----------------------------------------Regional-----------------------------------------
SELECT 
  
    RegionalName,   
    ISNULL(SUM(OpenNPPaktif), 0) AS OpenNPPaktif,
    ISNULL(SUM(OpenNPPod), 0) AS OpenNPPod,
    ISNULL(ROUND(CAST(ISNULL(SUM(OpenNPPod), 0) AS FLOAT) / NULLIF(CAST(SUM(OpenNPPaktif) AS FLOAT), 0) * 100, 2), 0) AS OpenPersenNPPOD,
    ISNULL(SUM(CloseNPPaktif), 0) AS CloseNPPaktif,
    ISNULL(SUM(CloseNPPod), 0) AS CloseNPPod,
    ISNULL(ROUND(CAST(ISNULL(SUM(CloseNPPod), 0) AS FLOAT) / NULLIF(CAST(SUM(CloseNPPaktif) AS FLOAT), 0) * 100, 2), 0) AS ClosePersenNPPOD,
    ISNULL(SUM(LastMonthOpenNPPaktif), 0) AS LastMonthOpenNPPaktif,
    ISNULL(SUM(LastMonthOpenNPPod), 0) AS LastMonthOpenNPPod,
    ISNULL(ROUND(CAST(ISNULL(SUM(LastMonthOpenNPPod), 0) AS FLOAT) / NULLIF(CAST(SUM(LastMonthOpenNPPaktif) AS FLOAT), 0) * 100, 2), 0) AS LastMonthOpenPersenNPPOD,
    ISNULL(SUM(LastMonthCloseNPPaktif), 0) AS LastMonthCloseNPPaktif,
    ISNULL(SUM(LastMonthCloseNPPod), 0) AS LastMonthCloseNPPod,
    ISNULL(ROUND(CAST(ISNULL(SUM(LastMonthCloseNPPod), 0) AS FLOAT) / NULLIF(CAST(SUM(LastMonthCloseNPPaktif) AS FLOAT), 0) * 100, 2), 0) AS LastMonthClosePersenNPPOD
INTO #regional
FROM #temp
--where RegionalName in (select RegionalName from #temp where konsolbranchid = @konsolbranchid)
Group by RegionalName

--select * from #regional
--drop table  #regional


	--select * from #m where wilayahname in (select wilayahname from #m where branchid = '456' )
-------------------------------------------------------area--------------------------------------
SELECT 

    areaname,
    ISNULL(SUM(OpenNPPaktif), 0) AS OpenNPPaktif,
    ISNULL(SUM(OpenNPPod), 0) AS OpenNPPod,
    ISNULL(ROUND(CAST(ISNULL(SUM(OpenNPPod), 0) AS FLOAT) / NULLIF(CAST(SUM(OpenNPPaktif) AS FLOAT), 0) * 100, 2), 0) AS OpenPersenNPPOD,
    ISNULL(SUM(CloseNPPaktif), 0) AS CloseNPPaktif,
    ISNULL(SUM(CloseNPPod), 0) AS CloseNPPod,
    ISNULL(ROUND(CAST(ISNULL(SUM(CloseNPPod), 0) AS FLOAT) / NULLIF(CAST(SUM(CloseNPPaktif) AS FLOAT), 0) * 100, 2), 0) AS ClosePersenNPPOD,
    ISNULL(SUM(LastMonthOpenNPPaktif), 0) AS LastMonthOpenNPPaktif,
    ISNULL(SUM(LastMonthOpenNPPod), 0) AS LastMonthOpenNPPod,
    ISNULL(ROUND(CAST(ISNULL(SUM(LastMonthOpenNPPod), 0) AS FLOAT) / NULLIF(CAST(SUM(LastMonthOpenNPPaktif) AS FLOAT), 0) * 100, 2), 0) AS LastMonthOpenPersenNPPOD,
    ISNULL(SUM(LastMonthCloseNPPaktif), 0) AS LastMonthCloseNPPaktif,
    ISNULL(SUM(LastMonthCloseNPPod), 0) AS LastMonthCloseNPPod,
    ISNULL(ROUND(CAST(ISNULL(SUM(LastMonthCloseNPPod), 0) AS FLOAT) / NULLIF(CAST(SUM(LastMonthCloseNPPaktif) AS FLOAT), 0) * 100, 2), 0) AS LastMonthClosePersenNPPOD
INTO #area
FROM #temp
--where AreaName in (select AreaName from #temp where konsolbranchid = @konsolbranchid )
GROUP BY areaname

--select * from #area
--drop table #area


--------------------------------------Jaringan--------------------------------------------------------------------------------
SELECT 

         Jaringanname,
         OpenNPPaktif = ISNULL(SUM(OpenNPPaktif), 0),
         OpenNPPod = ISNULL(SUM(OpenNPPod), 0),
         OpenPersenNPPOD =  ISNULL(ROUND(CAST(ISNULL(SUM(OpenNPPod), 0) AS float) / CAST(NULLIF(SUM(OpenNPPaktif), 0) AS float) * 100, 2),0),
         CloseNPPaktif = ISNULL(SUM(CloseNPPaktif), 0),
         CloseNPPod = ISNULL(SUM(CloseNPPod), 0),
         ClosePersenNPPOD =  ISNULL(ROUND(CAST(ISNULL(SUM(CloseNPPod), 0) AS float) / CAST(NULLIF(SUM(CloseNPPaktif), 0) AS float) * 100, 2),0),
         LastMonthOpenNPPaktif = ISNULL(SUM(LastMonthOpenNPPaktif), 0),
         LastMonthOpenNPPod = ISNULL(SUM(LastMonthOpenNPPod), 0),
         LastMonthOpenPersenNPPOD = ISNULL(ROUND(CAST(ISNULL(SUM(LastMonthOpenNPPod), 0) AS float) / CAST(NULLIF(SUM(LastMonthOpenNPPaktif), 0) AS float) * 100, 2),0),
         LastMonthCloseNPPaktif = ISNULL(SUM(LastMonthCloseNPPaktif), 0),
         LastMonthCloseNPPod = ISNULL(SUM(LastMonthCloseNPPod), 0),
         LastMonthClosePersenNPPOD = ISNULL(ROUND(CAST(ISNULL(SUM(LastMonthCloseNPPod), 0) AS float) / CAST(NULLIF(SUM(LastMonthCloseNPPaktif), 0) AS float) * 100, 2),0)
		 INTO #ODParentBranch
		 FROM #temp
		 --where Jaringanname in (select Jaringanname from #temp where konsolbranchid = @konsolbranchid )
         GROUP BY  Jaringanname

--select * from #ODParentBranch


--select * from #ODParentBranch
--drop table #ODParentBranch

----------------------------------- cabang pos*----------------------------------------------------------------------------------------

SELECT 

		 Branchname ,
         OpenNPPaktif = ISNULL(SUM(OpenNPPaktif), 0),
         OpenNPPod = ISNULL(SUM(OpenNPPod), 0),
         OpenPersenNPPOD =  ISNULL(ROUND(CAST(ISNULL(SUM(OpenNPPod), 0) AS float) / CAST(NULLIF(SUM(OpenNPPaktif), 0) AS float) * 100, 2),0),
         CloseNPPaktif = ISNULL(SUM(CloseNPPaktif), 0),
         CloseNPPod = ISNULL(SUM(CloseNPPod), 0),
         ClosePersenNPPOD =  ISNULL(ROUND(CAST(ISNULL(SUM(CloseNPPod), 0) AS float) / CAST(NULLIF(SUM(CloseNPPaktif), 0) AS float) * 100, 2),0),
         LastMonthOpenNPPaktif = ISNULL(SUM(LastMonthOpenNPPaktif), 0),
         LastMonthOpenNPPod = ISNULL(SUM(LastMonthOpenNPPod), 0),
         LastMonthOpenPersenNPPOD = ISNULL(ROUND(CAST(ISNULL(SUM(LastMonthOpenNPPod), 0) AS float) / CAST(NULLIF(SUM(LastMonthOpenNPPaktif), 0) AS float) * 100, 2),0),
         LastMonthCloseNPPaktif = ISNULL(SUM(LastMonthCloseNPPaktif), 0),
         LastMonthCloseNPPod = ISNULL(SUM(LastMonthCloseNPPod), 0),
         LastMonthClosePersenNPPOD = ISNULL(ROUND(CAST(ISNULL(SUM(LastMonthCloseNPPod), 0) AS float) / CAST(NULLIF(SUM(LastMonthCloseNPPaktif), 0) AS float) * 100, 2),0)
		 INTO #ODcabangpos
		 FROM #temp
		 --where Branchname in (select Branchname from #temp where konsolbranchid = @konsolbranchid )
	     GROUP BY Branchname

		
--select * from #nh
--select * from #ODcabangpos
---drop table #ODcabangpos
----------------------------------------head-----------------------------------------------------------------------------------------------------
    set @Pemisah=char(13)+char(10)
	set @MsgBody1 = ''
	set @MsgBody2 = ''
	set @MsgBody3 = ''
	set @MsgBody4 = ''
	set @MsgBody5 = '' 

	set @MsgBody6 = '' 
set @MsgBody1 = '' 
set @MsgHdr1 ='<html>
               <head>
                    <style type="text/css">
                        .table{
                            background-color: #F5F5F5;
                            border-collapse: collapse;
                            font-size: 12px;
                            font-family: "Verdana";}

                        .table .title{
                            text-align: center;
                            background-color: #3399FF;
          color:white;
                            font-weight:bold;}

                        .table .head{
                            text-align: center;
                            background-color: #FFCC00;}
						
                        .table td{
                            padding : 5px;
                            border  : solid 1px #555;
							font-weight:normal;}
                        
						.table .bad{
						    text-align: center;
                            background-color: #DC143C;}
						
						.table .good{
						    text-align: center;
                            background-color: #7FFF00;}
												
						.table .important{
						    text-align: center;
                            background-color: #Cad3e7;}
        </style>
		         </head>
				 <body>
	            <!-- TABLE PT -->
					</table>
					<br>
                    <table ID="PT" class="table">
					     <tr>
						   <td colspan="15" class="title" align=center> <b> Unit Over Due By PT </b> </td>
						</tr>
						<tr>
							<td rowspan="3" class="head"> <b> PT </b> </td>   
							<td colspan="7" class="head"> <b>' + convert(varchar(20),convert(datetime,@mtdThisMonth),106) + '</b> </td>
							<td colspan="7" class="head"> <b>' + convert(varchar(20),convert(datetime,@mtdlastMonth),106) + '</b> </td>							
						</tr>
						<tr>
							<td colspan="3" class="head"> <b> Open </b> </td>
							<td colspan="3" class="head"> <b> Close </b> </td>
							<td rowspan="2" class="head"> <b> Changes </b> </td>
							<td colspan="3" class="head"> <b> Open </b> </td>
							<td colspan="3" class="head"> <b> Close </b> </td>
							<td rowspan="2" class="head"> <b> Changes </b> </td>

						</tr>
						<tr>
							<td class="head"> <b> Unit Active </b> </td>
							<td class="head"> <b> Unit OD     </b> </td>  
							<td class="head"> <b> % OD        </b> </td>    
							<td class="head"> <b> Unit Active </b> </td>
							<td class="head"> <b> Unit OD     </b> </td>  
							<td class="head"> <b> % OD        </b> </td>  		
							<td class="head"> <b> Unit Active </b> </td>
							<td class="head"> <b> Unit OD     </b> </td>  
							<td class="head"> <b> % OD        </b> </td>    
							<td class="head"> <b> Unit Active </b> </td>
							<td class="head"> <b> Unit OD     </b> </td>  
							<td class="head"> <b> % OD        </b> </td>  											
						</tr>
            '

DECLARE cur1 CURSOR FOR 
SELECT * FROM #tempPT  ;

OPEN cur1;

FETCH NEXT FROM cur1 INTO @idx, @PT, @OpenNPPaktif,@OpenNPPod,@OpenPersenNPPOD,@CloseNPPaktif,@CloseNPPod,@ClosePersenNPPOD,@LastMonthOpenNPPaktif,
@LastMonthOpenNPPod,@LastMonthOpenPersenNPPOD,@LastMonthCloseNPPaktif,@LastMonthCloseNPPod,@LastMonthClosePersenNPPOD 
WHILE @@FETCH_STATUS = 0																		   
BEGIN


 SELECT 
        @idx = idx,
        @PT = PT,
        @OpenNPPAktif = OpenNPPAktif,
        @OpenNPPOD = OpenNPPOD,
        @OpenPersenNPPOD = OpenPersenNPPOD,
        @CloseNPPAktif = CloseNPPAktif,
        @CloseNPPOD = CloseNPPOD,
        @ClosePersenNPPOD = ClosePersenNPPOD,
        @LastMonthOpenNPPAktif = LastMonthOpenNPPAktif,
        @LastMonthOpenNPPOD = LastMonthOpenNPPOD,
        @LastMonthOpenPersenNPPOD = LastMonthOpenPersenNPPOD,
        @LastMonthCloseNPPAktif = LastMonthCloseNPPAktif,
        @LastMonthCloseNPPOD = LastMonthCloseNPPOD,
        @LastMonthClosePersenNPPOD = LastMonthClosePersenNPPOD
        from #tempPT
	    where idx=@CurrRec
	  

  		
		

SET @MsgBody1 = @MsgBody1 + 
					'
						<tr>
						<td class="title" align=center> <b>' +@PT+ '</b> </td>		
						<td class="td" align=center> ' +iif(@OpenNPPAktif=0,'0',convert(varchar(18),format(isnull(@OpenNPPAktif,0),'###,###')))+ '</td>		
						<td class="td" align=center> ' +iif(@OpenNPPOD=0,'0',convert(varchar(18),format(isnull(@OpenNPPOD,0),'###,###')))+ '</td>		
						<td class="important" align=center> <b>' +iif(@OpenPersenNPPOD=0,'0',convert(varchar(18),isnull(@OpenPersenNPPOD,0)))+ ' %</b> </td>	
						<td class="td" align=center> ' +iif(@CloseNPPAktif=0,'0',convert(varchar(18),format(isnull(@CloseNPPAktif,0),'###,###')))+ '</td>		
						<td class="td" align=center> ' +iif(@CloseNPPOD=0,'0',convert(varchar(18),format(isnull(@CloseNPPOD,0),'###,###')))+ '</td>		
						<td class="important" align=center> <b>' +iif(@ClosePersenNPPOD=0,'0',convert(varchar(18),isnull(@ClosePersenNPPOD,0)))+ ' %</b> </td>	
						<td class="important" align=center> <b>' +iif(@ClosePersenNPPOD-@OpenPersenNPPOD=0,'0',convert(varchar(18),isnull(@ClosePersenNPPOD-@OpenPersenNPPOD,0)))+ ' %</b> </td>	
						<td class="td" align=center>' +iif(@LastMonthOpenNPPAktif=0,'0',convert(varchar(18),format(isnull(@LastMonthOpenNPPAktif,0),'###,###')))+ '</td>		
						<td class="td" align=center>' +iif(@LastMonthOpenNPPOD=0,'0',convert(varchar(18),format(isnull(@LastMonthOpenNPPOD,0),'###,###')))+ '</td>		
						<td class="important" align=center> <b>' +iif(@LastMonthOpenPersenNPPOD=0,'0',convert(varchar(18),isnull(@LastMonthOpenPersenNPPOD,0)))+ ' %</b> </td>	
						<td class="td" align=center> ' +iif(@LastMonthCloseNPPAktif=0,'0',convert(varchar(18),format(isnull(@LastMonthCloseNPPAktif,0),'###,###')))+ '</td>		
						<td class="td" align=center> ' +iif(@LastMonthCloseNPPOD=0,'0',convert(varchar(18),format(isnull(@LastMonthCloseNPPOD,0),'###,###')))+ '</td>		
						<td class="important" align=center> <b>' +iif(@LastMonthClosePersenNPPOD=0,'0',convert(varchar(18),isnull(@LastMonthClosePersenNPPOD,0)))+ ' %</b> </td>	
						<td class="important" align=center> <b>' +iif(@LastMonthClosePersenNPPOD-@LastMonthOpenPersenNPPOD=0,'0',convert(varchar(18),isnull(@LastMonthClosePersenNPPOD-@LastMonthOpenPersenNPPOD,0)))+ ' %</b> </td>
						</tr>
					'
					SET @CurrRec=@CurrRec+1

FETCH NEXT FROM cur1 INTO @idx, @PT, @OpenNPPaktif,@OpenNPPod,@OpenPersenNPPOD,@CloseNPPaktif,@CloseNPPod,@ClosePersenNPPOD,@LastMonthOpenNPPaktif,
@LastMonthOpenNPPod,@LastMonthOpenPersenNPPOD,@LastMonthCloseNPPaktif,@LastMonthCloseNPPod,@LastMonthClosePersenNPPOD 
END


				select @TotalOpenAktifPT=sum(OpenNPPAktif) from #tempPT
				select @TotalOpenODPT=sum(OpenNPPOd) from #tempPT
				set @TotalOpenPersenODPT = round(CAST(@TotalOpenODPT as float)/CAST(@TotalOpenAktifPT as float)*100,2)

				select @TotalCloseAktifPT=sum(CloseNPPAktif) from #tempPT
				select @TotalCloseODPT=sum(CloseNPPOd) from #tempPT
				set @TotalClosePersenODPT = round(CAST(@TotalCloseODPT as float)/CAST(@TotalCloseAktifPT as float)*100,2)
			
				select @TotalLastMonthOpenAktifPT=sum(LastMonthOpenNPPAktif) from #tempPT
				select @TotalLastMonthOpenODPT=sum(LastMonthOpenNPPOd) from #tempPT
				set @TotalLastMonthOpenPersenODPT = round(CAST(@TotalLastMonthOpenODPT as float)/CAST(@TotalLastMonthOpenAktifPT as float)*100,2)

				select @TotalLastMonthCloseAktifPT=sum(LastMonthCloseNPPAktif) from #tempPT
				select @TotalLastMonthCloseODPT=sum(LastMonthCloseNPPOd) from #tempPT
				set @TotalLastMonthClosePersenODPT = round(CAST(@TotalLastMonthCloseODPT as float)/CAST(@TotalLastMonthCloseAktifPT as float)*100,2)


				SET @MsgBody1 = @MsgBody1 + 
						'
							<tr>
							<td class="head" align=center> <b> Total </b> </td>	
							<td class="head" align=center> ' +iif(@TotalOpenAktifPT=0,'0',convert(varchar(18),format(isnull(@TotalOpenAktifPT,0),'###,###')))+ '</td>		
							<td class="head" align=center> ' +iif(@TotalOpenODPT=0,'0',convert(varchar(18),format(isnull(@TotalOpenODPT,0),'###,###')))+ '</td>		
							<td class="head" align=center> <b>' +iif(@TotalOpenPersenODPT=0,'0',convert(varchar(18),isnull(@TotalOpenPersenODPT,0)))+ ' %</b> </td>	
							<td class="head" align=center> ' +iif(@TotalCloseAktifPT=0,'0',convert(varchar(18),format(isnull(@TotalCloseAktifPT,0),'###,###')))+ '</td>		
							<td class="head" align=center> ' +iif(@TotalCloseODPT=0,'0',convert(varchar(18),format(isnull(@TotalCloseODPT,0),'###,###')))+ '</td>		
							<td class="head" align=center> <b>' +iif(@TotalClosePersenODPT=0,'0',convert(varchar(18),isnull(@TotalClosePersenODPT,0)))+ ' %</b> </td>	
							<td class="head" align=center> <b>' +iif(@TotalClosePersenODPT-@TotalOpenPersenODPT=0,'0',convert(varchar(18),isnull(@TotalClosePersenODPT-@TotalOpenPersenODPT,0)))+ ' %</b> </td>
							<td class="head" align=center> ' +iif(@TotalLastMonthOpenAktifPT=0,'0',convert(varchar(18),format(isnull(@TotalLastMonthOpenAktifPT,0),'###,###')))+ '</td>		
							<td class="head" align=center> ' +iif(@TotalLastMonthOpenODPT=0,'0',convert(varchar(18),format(isnull(@TotalLastMonthOpenODPT,0),'###,###')))+ '</td>		
							<td class="head" align=center> <b>' +iif(@TotalLastMonthOpenPersenODPT=0,'0',convert(varchar(18),isnull(@TotalLastMonthOpenPersenODPT,0)))+ ' %</b> </td>	
							<td class="head" align=center> ' +iif(@TotalLastMonthCloseAktifPT=0,'0',convert(varchar(18),format(isnull(@TotalLastMonthCloseAktifPT,0),'###,###')))+ '</td>		
							<td class="head" align=center> ' +iif(@TotalLastMonthCloseODPT=0,'0',convert(varchar(18),format(isnull(@TotalLastMonthCloseODPT,0),'###,###')))+ '</td>		
							<td class="head" align=center> <b>' +iif(@TotalLastMonthClosePersenODPT=0,'0',convert(varchar(18),isnull(@TotalLastMonthClosePersenODPT,0)))+ ' %</b> </td>	
							<td class="head" align=center> <b>' +iif(@TotalLastMonthClosePersenODPT-@TotalLastMonthOpenPersenODPT=0,'0',convert(varchar(18),isnull(@TotalLastMonthClosePersenODPT-@TotalLastMonthOpenPersenODPT,0)))+ ' %</b> </td>
							</tr>
						'
close cur1
deallocate cur1


set @MsgBody2 = '' 
set @MsgHdr2 = '  </style>    
		           
	            <!-- TABLE Wilayah -->
					</table>
					<br>
                    <table ID="Wilayah" class="table">
					     <tr>
						   <td colspan="15" class="title" align=center> <b> Unit Over Due By Wilayah </b> </td>
						</tr>
						<tr>
							<td rowspan="3" class="head"> <b> Wilayah </b> </td>   
							<td colspan="7" class="head"> <b>' + convert(varchar(20),convert(datetime,@mtdThisMonth),106) + '</b> </td>
							<td colspan="7" class="head"> <b>' + convert(varchar(20),convert(datetime,@mtdlastMonth),106) + '</b> </td>							
						</tr>
						<tr>
							<td colspan="3" class="head"> <b> Open </b> </td>
							<td colspan="3" class="head"> <b> Close </b> </td>
							<td rowspan="2" class="head"> <b> Changes </b> </td>
							<td colspan="3" class="head"> <b> Open </b> </td>
							<td colspan="3" class="head"> <b> Close </b> </td>
							<td rowspan="2" class="head"> <b> Changes </b> </td>

						</tr>
						<tr>
							<td class="head"> <b> Unit Active </b> </td>
							<td class="head"> <b> Unit OD     </b> </td>  
							<td class="head"> <b> % OD        </b> </td>    
							<td class="head"> <b> Unit Active </b> </td>
							<td class="head"> <b> Unit OD     </b> </td>  
							<td class="head"> <b> % OD        </b> </td>  		
							<td class="head"> <b> Unit Active </b> </td>
							<td class="head"> <b> Unit OD     </b> </td>  
							<td class="head"> <b> % OD        </b> </td>    
							<td class="head"> <b> Unit Active </b> </td>
							<td class="head"> <b> Unit OD     </b> </td>  
							<td class="head"> <b> % OD        </b> </td>  											
						</tr>
            '




 SELECT 
        @HBAName =HBAname,
        @OpenNPPAktif = OpenNPPAktif,
        @OpenNPPOD = OpenNPPOD,
        @OpenPersenNPPOD = OpenPersenNPPOD,
        @CloseNPPAktif = CloseNPPAktif,
        @CloseNPPOD = CloseNPPOD,
        @ClosePersenNPPOD = ClosePersenNPPOD,
        @LastMonthOpenNPPAktif = LastMonthOpenNPPAktif,
        @LastMonthOpenNPPOD = LastMonthOpenNPPOD,
        @LastMonthOpenPersenNPPOD = LastMonthOpenPersenNPPOD,
        @LastMonthCloseNPPAktif = LastMonthCloseNPPAktif,
        @LastMonthCloseNPPOD = LastMonthCloseNPPOD,
        @LastMonthClosePersenNPPOD = LastMonthClosePersenNPPOD
        from #wilayah
	    where HBAname in (select HBAname from #temp where BranchID = @branchid )

SET @MsgBody2 = @MsgBody2 + 
					'
						<tr>
						<td class="title" align=center> <b>' +@HBAname+ '</b> </td>		
						<td class="td" align=center> ' +iif(@OpenNPPAktif=0,'0',convert(varchar(18),format(isnull(@OpenNPPAktif,0),'###,###')))+ '</td>		
						<td class="td" align=center> ' +iif(@OpenNPPOD=0,'0',convert(varchar(18),format(isnull(@OpenNPPOD,0),'###,###')))+ '</td>		
						<td class="important" align=center> <b>' +iif(@OpenPersenNPPOD=0,'0',convert(varchar(18),isnull(@OpenPersenNPPOD,0)))+ ' %</b> </td>	
						<td class="td" align=center> ' +iif(@CloseNPPAktif=0,'0',convert(varchar(18),format(isnull(@CloseNPPAktif,0),'###,###')))+ '</td>		
						<td class="td" align=center> ' +iif(@CloseNPPOD=0,'0',convert(varchar(18),format(isnull(@CloseNPPOD,0),'###,###')))+ '</td>		
						<td class="important" align=center> <b>' +iif(@ClosePersenNPPOD=0,'0',convert(varchar(18),isnull(@ClosePersenNPPOD,0)))+ ' %</b> </td>	
						<td class="important" align=center> <b>' +iif(@ClosePersenNPPOD-@OpenPersenNPPOD=0,'0',convert(varchar(18),isnull(@ClosePersenNPPOD-@OpenPersenNPPOD,0)))+ ' %</b> </td>	
						<td class="td" align=center>' +iif(@LastMonthOpenNPPAktif=0,'0',convert(varchar(18),format(isnull(@LastMonthOpenNPPAktif,0),'###,###')))+ '</td>		
						<td class="td" align=center>' +iif(@LastMonthOpenNPPOD=0,'0',convert(varchar(18),format(isnull(@LastMonthOpenNPPOD,0),'###,###')))+ '</td>		
						<td class="important" align=center> <b>' +iif(@LastMonthOpenPersenNPPOD=0,'0',convert(varchar(18),isnull(@LastMonthOpenPersenNPPOD,0)))+ ' %</b> </td>	
						<td class="td" align=center> ' +iif(@LastMonthCloseNPPAktif=0,'0',convert(varchar(18),format(isnull(@LastMonthCloseNPPAktif,0),'###,###')))+ '</td>		
						<td class="td" align=center> ' +iif(@LastMonthCloseNPPOD=0,'0',convert(varchar(18),format(isnull(@LastMonthCloseNPPOD,0),'###,###')))+ '</td>		
						<td class="important" align=center> <b>' +iif(@LastMonthClosePersenNPPOD=0,'0',convert(varchar(18),isnull(@LastMonthClosePersenNPPOD,0)))+ ' %</b> </td>	
						<td class="important" align=center> <b>' +iif(@LastMonthClosePersenNPPOD-@LastMonthOpenPersenNPPOD=0,'0',convert(varchar(18),isnull(@LastMonthClosePersenNPPOD-@LastMonthOpenPersenNPPOD,0)))+ ' %</b> </td>
						</tr>
					'
               

				select @TotalOpenAktifwilayah=sum(OpenNPPAktif) from #wilayah where HBAname in (select HBAname from #temp where BranchID = @branchid )
                select @TotalOpenODwilayah=sum(OpenNPPOd) from #wilayah where HBAname in (select HBAname from #temp where BranchID = @branchid )
                set @TotalOpenPersenODwilayah = round(CAST(@TotalOpenODwilayah as float)/CAST(@TotalOpenAktifwilayah as float)*100,2) 
                
                select @TotalCloseAktifwilayah=sum(CloseNPPAktif) from #wilayah where HBAname in (select HBAname from #temp where BranchID = @branchid )
                select @TotalCloseODwilayah=sum(CloseNPPOd) from #wilayah where HBAname in (select HBAname from #temp where BranchID = @branchid )
                set @TotalClosePersenODwilayah = round(CAST(@TotalCloseODwilayah as float)/CAST(@TotalCloseAktifwilayah as float)*100,2)
                
                select @TotalLastMonthOpenAktifwilayah=sum(LastMonthOpenNPPAktif) from #wilayah where HBAname in (select HBAname from #temp where BranchID = @branchid )
                select @TotalLastMonthOpenODwilayah=sum(LastMonthOpenNPPOd) from #wilayah where HBAname in (select HBAname from #temp where BranchID = @branchid )
                set @TotalLastMonthOpenPersenODwilayah = round(CAST(@TotalLastMonthOpenODwilayah as float)/CAST(@TotalLastMonthOpenAktifwilayah as float)*100,2)
                
                select @TotalLastMonthCloseAktifwilayah=sum(LastMonthCloseNPPAktif) from #wilayah where HBAname in (select HBAname from #temp where BranchID = @branchid )
                select @TotalLastMonthCloseODwilayah=sum(LastMonthCloseNPPOd) from #wilayah where HBAname in (select HBAname from #temp where BranchID = @branchid )
                set @TotalLastMonthClosePersenODwilayah = round(CAST(@TotalLastMonthCloseODwilayah as float)/CAST(@TotalLastMonthCloseAktifwilayah as float)*100,2)
                
				SET @MsgBody2 = @MsgBody2 + 
						'
							<tr>
							<td class="head" align=center> <b> Total </b> </td>	
							<td class="head" align=center> ' +iif(@TotalOpenAktifwilayah=0,'0',convert(varchar(18),format(isnull(@TotalOpenAktifwilayah,0),'###,###')))+ '</td>		
                            <td class="head" align=center> ' +iif(@TotalOpenODwilayah=0,'0',convert(varchar(18),format(isnull(@TotalOpenODwilayah,0),'###,###')))+ '</td>		
                            <td class="head" align=center> <b>' +iif(@TotalOpenPersenODwilayah=0,'0',convert(varchar(18),isnull(@TotalOpenPersenODwilayah,0)))+ ' %</b> </td>	
                            <td class="head" align=center> ' +iif(@TotalCloseAktifwilayah=0,'0',convert(varchar(18),format(isnull(@TotalCloseAktifwilayah,0),'###,###')))+ '</td>		
                            <td class="head" align=center> ' +iif(@TotalCloseODwilayah=0,'0',convert(varchar(18),format(isnull(@TotalCloseODwilayah,0),'###,###')))+ '</td>		
                            <td class="head" align=center> <b>' +iif(@TotalClosePersenODwilayah=0,'0',convert(varchar(18),isnull(@TotalClosePersenODwilayah,0)))+ ' %</b> </td>	
                            <td class="head" align=center> <b>' +iif(@TotalClosePersenODwilayah-@TotalOpenPersenODwilayah=0,'0',convert(varchar(18),isnull(@TotalClosePersenODwilayah-@TotalOpenPersenODwilayah,0)))+ ' %</b> </td>
                            <td class="head" align=center> ' +iif(@TotalLastMonthOpenAktifwilayah=0,'0',convert(varchar(18),format(isnull(@TotalLastMonthOpenAktifwilayah,0),'###,###')))+ '</td>		
                            <td class="head" align=center> ' +iif(@TotalLastMonthOpenODwilayah=0,'0',convert(varchar(18),format(isnull(@TotalLastMonthOpenODwilayah,0),'###,###')))+ '</td>		
                            <td class="head" align=center> <b>' +iif(@TotalLastMonthOpenPersenODwilayah=0,'0',convert(varchar(18),isnull(@TotalLastMonthOpenPersenODwilayah,0)))+ ' %</b> </td>	
                            <td class="head" align=center> ' +iif(@TotalLastMonthCloseAktifwilayah=0,'0',convert(varchar(18),format(isnull(@TotalLastMonthCloseAktifwilayah,0),'###,###')))+ '</td>		
                            <td class="head" align=center> ' +iif(@TotalLastMonthCloseODwilayah=0,'0',convert(varchar(18),format(isnull(@TotalLastMonthCloseODwilayah,0),'###,###')))+ '</td>		
                            <td class="head" align=center> <b>' +iif(@TotalLastMonthClosePersenODwilayah=0,'0',convert(varchar(18),isnull(@TotalLastMonthClosePersenODwilayah,0)))+ ' %</b> </td>	
                            <td class="head" align=center> <b>' +iif(@TotalLastMonthClosePersenODwilayah-@TotalLastMonthOpenPersenODwilayah=0,'0',convert(varchar(18),isnull(@TotalLastMonthClosePersenODwilayah-@TotalLastMonthOpenPersenODwilayah,0)))+ ' %</b> </td>																					
							</tr>
						'


set @MsgBody3 = '' 
set @MsgHdr3= '  </style>    
		           
	            <!-- TABLE Region -->
					</table>
					<br>
                    <table ID="Region" class="table">
					     <tr>
						   <td colspan="15" class="title" align=center> <b> Unit Over Due By Region </b> </td>
						</tr>
						<tr>
							<td rowspan="3" class="head"> <b> Region </b> </td>   
							<td colspan="7" class="head"> <b>' + convert(varchar(20),convert(datetime,@mtdThisMonth),106) + '</b> </td>
							<td colspan="7" class="head"> <b>' + convert(varchar(20),convert(datetime,@mtdlastMonth),106) + '</b> </td>							
						</tr>
						<tr>
							<td colspan="3" class="head"> <b> Open </b> </td>
							<td colspan="3" class="head"> <b> Close </b> </td>
							<td rowspan="2" class="head"> <b> Changes </b> </td>
							<td colspan="3" class="head"> <b> Open </b> </td>
							<td colspan="3" class="head"> <b> Close </b> </td>
							<td rowspan="2" class="head"> <b> Changes </b> </td>

						</tr>
						<tr>
							<td class="head"> <b> Unit Active </b> </td>
							<td class="head"> <b> Unit OD     </b> </td>  
							<td class="head"> <b> % OD        </b> </td>    
							<td class="head"> <b> Unit Active </b> </td>
							<td class="head"> <b> Unit OD     </b> </td>  
							<td class="head"> <b> % OD        </b> </td>  		
							<td class="head"> <b> Unit Active </b> </td>
							<td class="head"> <b> Unit OD     </b> </td>  
							<td class="head"> <b> % OD        </b> </td>    
							<td class="head"> <b> Unit Active </b> </td>
							<td class="head"> <b> Unit OD     </b> </td>  
							<td class="head"> <b> % OD        </b> </td>  											
						</tr>
            '




SELECT 
	
        @regionalname = regionalname,
        @OpenNPPAktif = OpenNPPAktif,
        @OpenNPPOD = OpenNPPOD,
        @OpenPersenNPPOD = OpenPersenNPPOD,
        @CloseNPPAktif = CloseNPPAktif,
        @CloseNPPOD = CloseNPPOD,
        @ClosePersenNPPOD = ClosePersenNPPOD,
        @LastMonthOpenNPPAktif = LastMonthOpenNPPAktif,
        @LastMonthOpenNPPOD = LastMonthOpenNPPOD,
        @LastMonthOpenPersenNPPOD = LastMonthOpenPersenNPPOD,
        @LastMonthCloseNPPAktif = LastMonthCloseNPPAktif,
        @LastMonthCloseNPPOD = LastMonthCloseNPPOD,
        @LastMonthClosePersenNPPOD = LastMonthClosePersenNPPOD
        from  #regional
		where regionalname in (select regionalname from #temp where BranchID = @branchid)


SET @MsgBody3 = @MsgBody3 + 
					'
						<tr>
						<td class="title" align=center> <b>' +@regionalname+ '</b> </td>		
						<td class="td" align=center> ' +iif(@OpenNPPAktif=0,'0',convert(varchar(18),format(isnull(@OpenNPPAktif,0),'###,###')))+ '</td>		
						<td class="td" align=center> ' +iif(@OpenNPPOD=0,'0',convert(varchar(18),format(isnull(@OpenNPPOD,0),'###,###')))+ '</td>		
						<td class="important" align=center> <b>' +iif(@OpenPersenNPPOD=0,'0',convert(varchar(18),isnull(@OpenPersenNPPOD,0)))+ ' %</b> </td>	
						<td class="td" align=center> ' +iif(@CloseNPPAktif=0,'0',convert(varchar(18),format(isnull(@CloseNPPAktif,0),'###,###')))+ '</td>		
						<td class="td" align=center> ' +iif(@CloseNPPOD=0,'0',convert(varchar(18),format(isnull(@CloseNPPOD,0),'###,###')))+ '</td>		
						<td class="important" align=center> <b>' +iif(@ClosePersenNPPOD=0,'0',convert(varchar(18),isnull(@ClosePersenNPPOD,0)))+ ' %</b> </td>	
						<td class="important" align=center> <b>' +iif(@ClosePersenNPPOD-@OpenPersenNPPOD=0,'0',convert(varchar(18),isnull(@ClosePersenNPPOD-@OpenPersenNPPOD,0)))+ ' %</b> </td>	
						<td class="td" align=center>' +iif(@LastMonthOpenNPPAktif=0,'0',convert(varchar(18),format(isnull(@LastMonthOpenNPPAktif,0),'###,###')))+ '</td>		
						<td class="td" align=center>' +iif(@LastMonthOpenNPPOD=0,'0',convert(varchar(18),format(isnull(@LastMonthOpenNPPOD,0),'###,###')))+ '</td>		
						<td class="important" align=center> <b>' +iif(@LastMonthOpenPersenNPPOD=0,'0',convert(varchar(18),isnull(@LastMonthOpenPersenNPPOD,0)))+ ' %</b> </td>	
						<td class="td" align=center> ' +iif(@LastMonthCloseNPPAktif=0,'0',convert(varchar(18),format(isnull(@LastMonthCloseNPPAktif,0),'###,###')))+ '</td>		
						<td class="td" align=center> ' +iif(@LastMonthCloseNPPOD=0,'0',convert(varchar(18),format(isnull(@LastMonthCloseNPPOD,0),'###,###')))+ '</td>		
						<td class="important" align=center> <b>' +iif(@LastMonthClosePersenNPPOD=0,'0',convert(varchar(18),isnull(@LastMonthClosePersenNPPOD,0)))+ ' %</b> </td>	
						<td class="important" align=center> <b>' +iif(@LastMonthClosePersenNPPOD-@LastMonthOpenPersenNPPOD=0,'0',convert(varchar(18),isnull(@LastMonthClosePersenNPPOD-@LastMonthOpenPersenNPPOD,0)))+ ' %</b> </td>
						</tr>
					'
		
				

				select @TotalOpenAktifRegional=sum(OpenNPPAktif) from #regional where regionalname in (select regionalname from #temp where BranchID = @branchid)
                select @TotalOpenODRegional=sum(OpenNPPOd) from #regional where regionalname in (select regionalname from #temp where BranchID = @branchid)
                set @TotalOpenPersenODRegional = round(CAST(@TotalOpenODRegional as float)/CAST(@TotalOpenAktifRegional as float)*100,2)
                
                select @TotalCloseAktifRegional=sum(CloseNPPAktif) from #regional where regionalname in (select regionalname from #temp where BranchID = @branchid)
                select @TotalCloseODRegional=sum(CloseNPPOd) from #regional where regionalname in (select regionalname from #temp where BranchID = @branchid)
                set @TotalClosePersenODRegional = round(CAST(@TotalCloseODRegional as float)/CAST(@TotalCloseAktifRegional as float)*100,2)
                
                select @TotalLastMonthOpenAktifRegional=sum(LastMonthOpenNPPAktif) from #regional where regionalname in (select regionalname from #temp where BranchID = @branchid)
                select @TotalLastMonthOpenODRegional=sum(LastMonthOpenNPPOd) from #regional where regionalname in (select regionalname from #temp where BranchID = @branchid)
                set @TotalLastMonthOpenPersenODRegional = round(CAST(@TotalLastMonthOpenODRegional as float)/CAST(@TotalLastMonthOpenAktifRegional as float)*100,2)
                
                select @TotalLastMonthCloseAktifRegional=sum(LastMonthCloseNPPAktif) from #regional where regionalname in (select regionalname from #temp where BranchID = @branchid)
                select @TotalLastMonthCloseODRegional=sum(LastMonthCloseNPPOd) from #regional where regionalname in (select regionalname from #temp where BranchID = @branchid)
                set @TotalLastMonthClosePersenODRegional = round(CAST(@TotalLastMonthCloseODRegional as float)/CAST(@TotalLastMonthCloseAktifRegional as float)*100,2)
                											
				SET @MsgBody3 = @MsgBody3 + 
						'
							<tr>
							<td class="head" align=center> <b> Total </b> </td>	
                            <td class="head" align=center> ' +iif(@TotalOpenAktifRegional=0,'0',convert(varchar(18),format(isnull(@TotalOpenAktifRegional,0),'###,###')))+ '</td>		
                            <td class="head" align=center> ' +iif(@TotalOpenODRegional=0,'0',convert(varchar(18),format(isnull(@TotalOpenODRegional,0),'###,###')))+ '</td>		
                            <td class="head" align=center> <b>' +iif(@TotalOpenPersenODRegional=0,'0',convert(varchar(18),isnull(@TotalOpenPersenODRegional,0)))+ ' %</b> </td>	
                            <td class="head" align=center> ' +iif(@TotalCloseAktifRegional=0,'0',convert(varchar(18),format(isnull(@TotalCloseAktifRegional,0),'###,###')))+ '</td>		
                            <td class="head" align=center> ' +iif(@TotalCloseODRegional=0,'0',convert(varchar(18),format(isnull(@TotalCloseODRegional,0),'###,###')))+ '</td>		
                            <td class="head" align=center> <b>' +iif(@TotalClosePersenODRegional=0,'0',convert(varchar(18),isnull(@TotalClosePersenODRegional,0)))+ ' %</b> </td>	
                            <td class="head" align=center> <b>' +iif(@TotalClosePersenODRegional-@TotalOpenPersenODRegional=0,'0',convert(varchar(18),isnull(@TotalClosePersenODRegional-@TotalOpenPersenODRegional,0)))+ ' %</b> </td>
                            <td class="head" align=center> ' +iif(@TotalLastMonthOpenAktifRegional=0,'0',convert(varchar(18),format(isnull(@TotalLastMonthOpenAktifRegional,0),'###,###')))+ '</td>		
                            <td class="head" align=center> ' +iif(@TotalLastMonthOpenODRegional=0,'0',convert(varchar(18),format(isnull(@TotalLastMonthOpenODRegional,0),'###,###')))+ '</td>		
                            <td class="head" align=center> <b>' +iif(@TotalLastMonthOpenPersenODRegional=0,'0',convert(varchar(18),isnull(@TotalLastMonthOpenPersenODRegional,0)))+ ' %</b> </td>	
                            <td class="head" align=center> ' +iif(@TotalLastMonthCloseAktifRegional=0,'0',convert(varchar(18),format(isnull(@TotalLastMonthCloseAktifRegional,0),'###,###')))+ '</td>		
                            <td class="head" align=center> ' +iif(@TotalLastMonthCloseODRegional=0,'0',convert(varchar(18),format(isnull(@TotalLastMonthCloseODRegional,0),'###,###')))+ '</td>		
                            <td class="head" align=center> <b>' +iif(@TotalLastMonthClosePersenODRegional=0,'0',convert(varchar(18),isnull(@TotalLastMonthClosePersenODRegional,0)))+ ' %</b> </td>	
                            <td class="head" align=center> <b>' +iif(@TotalLastMonthClosePersenODRegional-@TotalLastMonthOpenPersenODRegional=0,'0',convert(varchar(18),isnull(@TotalLastMonthClosePersenODRegional-@TotalLastMonthOpenPersenODRegional,0)))+ ' %</b> </td>														
							</tr>
						'


set @MsgBody4 = '' 
set @MsgHdr4= '  </style>    
		           
	            <!-- TABLE Area -->
					</table>
					<br>
                    <table ID="Area" class="table">
					     <tr>
						   <td colspan="15" class="title" align=center> <b> Unit Over Due By Area </b> </td>
						</tr>
						<tr>
							<td rowspan="3" class="head"> <b> Area </b> </td>   
							<td colspan="7" class="head"> <b>' + convert(varchar(20),convert(datetime,@mtdThisMonth),106) + '</b> </td>
							<td colspan="7" class="head"> <b>' + convert(varchar(20),convert(datetime,@mtdlastMonth),106) + '</b> </td>							
						</tr>
						<tr>
							<td colspan="3" class="head"> <b> Open </b> </td>
							<td colspan="3" class="head"> <b> Close </b> </td>
							<td rowspan="2" class="head"> <b> Changes </b> </td>
							<td colspan="3" class="head"> <b> Open </b> </td>
							<td colspan="3" class="head"> <b> Close </b> </td>
							<td rowspan="2" class="head"> <b> Changes </b> </td>

						</tr>
						<tr>
							<td class="head"> <b> Unit Active </b> </td>
							<td class="head"> <b> Unit OD     </b> </td>  
							<td class="head"> <b> % OD        </b> </td>    
							<td class="head"> <b> Unit Active </b> </td>
							<td class="head"> <b> Unit OD     </b> </td>  
							<td class="head"> <b> % OD        </b> </td>  		
							<td class="head"> <b> Unit Active </b> </td>
							<td class="head"> <b> Unit OD     </b> </td>  
							<td class="head"> <b> % OD        </b> </td>    
							<td class="head"> <b> Unit Active </b> </td>
							<td class="head"> <b> Unit OD     </b> </td>  
							<td class="head"> <b> % OD        </b> </td>  											
						</tr>
            '

--SELECT * FROM #area where areaname in (select areaname from #temp where BranchID = @branchid ) ;

 SELECT 

        @areaname = areaname,
        @OpenNPPAktif = OpenNPPAktif,
        @OpenNPPOD = OpenNPPOD,
        @OpenPersenNPPOD = OpenPersenNPPOD,
        @CloseNPPAktif = CloseNPPAktif,
        @CloseNPPOD = CloseNPPOD,
        @ClosePersenNPPOD = ClosePersenNPPOD,
        @LastMonthOpenNPPAktif = LastMonthOpenNPPAktif,
        @LastMonthOpenNPPOD = LastMonthOpenNPPOD,
        @LastMonthOpenPersenNPPOD = LastMonthOpenPersenNPPOD,
        @LastMonthCloseNPPAktif = LastMonthCloseNPPAktif,
        @LastMonthCloseNPPOD = LastMonthCloseNPPOD,
        @LastMonthClosePersenNPPOD = LastMonthClosePersenNPPOD
        from  #area
		where areaname in (select areaname from #temp where BranchID = @branchid )
	


SET @MsgBody4 = @MsgBody4 + 
					'
						<tr>
						<td class="title" align=center> <b>' +@areaname+ '</b> </td>		
						<td class="td" align=center> ' +iif(@OpenNPPAktif=0,'0',convert(varchar(18),format(isnull(@OpenNPPAktif,0),'###,###')))+ '</td>		
						<td class="td" align=center> ' +iif(@OpenNPPOD=0,'0',convert(varchar(18),format(isnull(@OpenNPPOD,0),'###,###')))+ '</td>		
						<td class="important" align=center> <b>' +iif(@OpenPersenNPPOD=0,'0',convert(varchar(18),isnull(@OpenPersenNPPOD,0)))+ ' %</b> </td>	
						<td class="td" align=center> ' +iif(@CloseNPPAktif=0,'0',convert(varchar(18),format(isnull(@CloseNPPAktif,0),'###,###')))+ '</td>		
						<td class="td" align=center> ' +iif(@CloseNPPOD=0,'0',convert(varchar(18),format(isnull(@CloseNPPOD,0),'###,###')))+ '</td>		
						<td class="important" align=center> <b>' +iif(@ClosePersenNPPOD=0,'0',convert(varchar(18),isnull(@ClosePersenNPPOD,0)))+ ' %</b> </td>	
						<td class="important" align=center> <b>' +iif(@ClosePersenNPPOD-@OpenPersenNPPOD=0,'0',convert(varchar(18),isnull(@ClosePersenNPPOD-@OpenPersenNPPOD,0)))+ ' %</b> </td>	
						<td class="td" align=center>' +iif(@LastMonthOpenNPPAktif=0,'0',convert(varchar(18),format(isnull(@LastMonthOpenNPPAktif,0),'###,###')))+ '</td>		
						<td class="td" align=center>' +iif(@LastMonthOpenNPPOD=0,'0',convert(varchar(18),format(isnull(@LastMonthOpenNPPOD,0),'###,###')))+ '</td>		
						<td class="important" align=center> <b>' +iif(@LastMonthOpenPersenNPPOD=0,'0',convert(varchar(18),isnull(@LastMonthOpenPersenNPPOD,0)))+ ' %</b> </td>	
						<td class="td" align=center> ' +iif(@LastMonthCloseNPPAktif=0,'0',convert(varchar(18),format(isnull(@LastMonthCloseNPPAktif,0),'###,###')))+ '</td>		
						<td class="td" align=center> ' +iif(@LastMonthCloseNPPOD=0,'0',convert(varchar(18),format(isnull(@LastMonthCloseNPPOD,0),'###,###')))+ '</td>		
						<td class="important" align=center> <b>' +iif(@LastMonthClosePersenNPPOD=0,'0',convert(varchar(18),isnull(@LastMonthClosePersenNPPOD,0)))+ ' %</b> </td>	
						<td class="important" align=center> <b>' +iif(@LastMonthClosePersenNPPOD-@LastMonthOpenPersenNPPOD=0,'0',convert(varchar(18),isnull(@LastMonthClosePersenNPPOD-@LastMonthOpenPersenNPPOD,0)))+ ' %</b> </td>
						</tr>
					'
			    
				
			    
SELECT @TotalOpenAktifarea = ISNULL(SUM(OpenNPPAktif), 0) FROM #area where areaname in (select areaname from #temp where BranchID = @branchid )
SELECT @TotalOpenODarea = ISNULL(SUM(OpenNPPOd), 0) FROM #area where areaname in (select areaname from #temp where BranchID = @branchid )
SET @TotalOpenPersenODarea = ISNULL(ROUND(CAST(@TotalOpenODarea AS FLOAT) /  NULLIF(CAST(@TotalOpenAktifarea AS FLOAT), 0) * 100, 2), 0)
SELECT @TotalCloseAktifarea = ISNULL(SUM(CloseNPPAktif), 0) FROM #area where areaname in (select areaname from #temp where BranchID = @branchid )
SELECT @TotalCloseODarea = ISNULL(SUM(CloseNPPOd), 0) FROM #area where areaname in (select areaname from #temp where BranchID = @branchid )
SET @TotalClosePersenODarea = ISNULL(ROUND(CAST(@TotalCloseODarea AS FLOAT) / NULLIF(CAST(@TotalCloseAktifarea AS FLOAT), 0) * 100, 2), 0)


SELECT @TotalLastMonthOpenAktifarea = ISNULL(SUM(LastMonthOpenNPPAktif), 0) FROM #area where areaname in (select areaname from #temp where BranchID = @branchid )
SELECT @TotalLastMonthOpenODarea = ISNULL(SUM(LastMonthOpenNPPOd), 0) FROM #area where areaname in (select areaname from #temp where BranchID = @branchid )
SET @TotalLastMonthOpenPersenODarea = ISNULL(ROUND(CAST(@TotalLastMonthOpenODarea AS FLOAT) / NULLIF(CAST(@TotalLastMonthOpenAktifarea AS FLOAT), 0) * 100, 2), 0)

SELECT @TotalLastMonthCloseAktifarea = ISNULL(SUM(LastMonthCloseNPPAktif), 0) FROM #area where areaname in (select areaname from #temp where BranchID = @branchid )
SELECT @TotalLastMonthCloseODarea = ISNULL(SUM(LastMonthCloseNPPOd), 0) FROM #area where areaname in (select areaname from #temp where BranchID = @branchid )
SET @TotalLastMonthClosePersenODarea = ISNULL(ROUND(CAST(@TotalLastMonthCloseODarea AS FLOAT) / NULLIF(CAST(@TotalLastMonthCloseAktifarea AS FLOAT), 0) * 100, 2), 0)
		
				SET @MsgBody4 = @MsgBody4 + 
						'
							<tr>
							<td class="head" align=center> <b> Total </b> </td>	
							<td class="head" align=center> ' +iif(@TotalOpenAktifarea=0,'0',convert(varchar(18),format(isnull(@TotalOpenAktifarea,0),'###,###')))+ '</td>		
                            <td class="head" align=center> ' +iif(@TotalOpenODarea=0,'0',convert(varchar(18),format(isnull(@TotalOpenODarea,0),'###,###')))+ '</td>		
                            <td class="head" align=center> <b>' +iif(@TotalOpenPersenODarea=0,'0',convert(varchar(18),isnull(@TotalOpenPersenODarea,0)))+ ' %</b> </td>	
                            <td class="head" align=center> ' +iif(@TotalCloseAktifarea=0,'0',convert(varchar(18),format(isnull(@TotalCloseAktifarea,0),'###,###')))+ '</td>		
                            <td class="head" align=center> ' +iif(@TotalCloseODarea=0,'0',convert(varchar(18),format(isnull(@TotalCloseODarea,0),'###,###')))+ '</td>		
                            <td class="head" align=center> <b>' +iif(@TotalClosePersenODarea=0,'0',convert(varchar(18),isnull(@TotalClosePersenODarea,0)))+ ' %</b> </td>	
                            <td class="head" align=center> <b>' +iif(@TotalClosePersenODarea-@TotalOpenPersenODarea=0,'0',convert(varchar(18),isnull(@TotalClosePersenODarea-@TotalOpenPersenODarea,0)))+ ' %</b> </td>
                            <td class="head" align=center> ' +iif(@TotalLastMonthOpenAktifarea=0,'0',convert(varchar(18),format(isnull(@TotalLastMonthOpenAktifarea,0),'###,###')))+ '</td>		
                            <td class="head" align=center> ' +iif(@TotalLastMonthOpenODarea=0,'0',convert(varchar(18),format(isnull(@TotalLastMonthOpenODarea,0),'###,###')))+ '</td>		
                            <td class="head" align=center> <b>' +iif(@TotalLastMonthOpenPersenODarea=0,'0',convert(varchar(18),isnull(@TotalLastMonthOpenPersenODarea,0)))+ ' %</b> </td>	
                            <td class="head" align=center> ' +iif(@TotalLastMonthCloseAktifarea=0,'0',convert(varchar(18),format(isnull(@TotalLastMonthCloseAktifarea,0),'###,###')))+ '</td>		
                            <td class="head" align=center> ' +iif(@TotalLastMonthCloseODarea=0,'0',convert(varchar(18),format(isnull(@TotalLastMonthCloseODarea,0),'###,###')))+ '</td>		
                            <td class="head" align=center> <b>' +iif(@TotalLastMonthClosePersenODarea=0,'0',convert(varchar(18),isnull(@TotalLastMonthClosePersenODarea,0)))+ ' %</b> </td>	
                            <td class="head" align=center> <b>' +iif(@TotalLastMonthClosePersenODarea-@TotalLastMonthOpenPersenODarea=0,'0',convert(varchar(18),isnull(@TotalLastMonthClosePersenODarea-@TotalLastMonthOpenPersenODarea,0)))+ ' %</b> </td>																				
							</tr>
						'

set @MsgBody5 = '' 
set @MsgHdr5 = '  </style>    
		           
	            <!-- TABLE Jaringan -->
					</table>
					<br>
                    <table ID="Jaringan" class="table">
					     <tr>
						   <td colspan="15" class="title" align=center> <b> Unit Over Due By Jaringan </b> </td>
						</tr>
						<tr>
							<td rowspan="3" class="head"> <b> Jaringan </b> </td>   
							<td colspan="7" class="head"> <b>' + convert(varchar(20),convert(datetime,@mtdThisMonth),106) + '</b> </td>
							<td colspan="7" class="head"> <b>' + convert(varchar(20),convert(datetime,@mtdlastMonth),106) + '</b> </td>							
						</tr>
						<tr>
							<td colspan="3" class="head"> <b> Open </b> </td>
							<td colspan="3" class="head"> <b> Close </b> </td>
							<td rowspan="2" class="head"> <b> Changes </b> </td>
							<td colspan="3" class="head"> <b> Open </b> </td>
							<td colspan="3" class="head"> <b> Close </b> </td>
							<td rowspan="2" class="head"> <b> Changes </b> </td>

						</tr>
						<tr>
							<td class="head"> <b> Unit Active </b> </td>
							<td class="head"> <b> Unit OD     </b> </td>  
							<td class="head"> <b> % OD        </b> </td>    
							<td class="head"> <b> Unit Active </b> </td>
							<td class="head"> <b> Unit OD     </b> </td>  
							<td class="head"> <b> % OD        </b> </td>  		
							<td class="head"> <b> Unit Active </b> </td>
							<td class="head"> <b> Unit OD     </b> </td>  
							<td class="head"> <b> % OD        </b> </td>    
							<td class="head"> <b> Unit Active </b> </td>
							<td class="head"> <b> Unit OD     </b> </td>  
							<td class="head"> <b> % OD        </b> </td>  											
						</tr>
            '



    SELECT 
        @parentbranch = Jaringanname,
        @OpenNPPAktif = OpenNPPAktif,
        @OpenNPPOD = OpenNPPOD,
        @OpenPersenNPPOD = OpenPersenNPPOD,
        @CloseNPPAktif = CloseNPPAktif,
        @CloseNPPOD = CloseNPPOD,
        @ClosePersenNPPOD = ClosePersenNPPOD,
        @LastMonthOpenNPPAktif = LastMonthOpenNPPAktif,
        @LastMonthOpenNPPOD = LastMonthOpenNPPOD,
        @LastMonthOpenPersenNPPOD = LastMonthOpenPersenNPPOD,
        @LastMonthCloseNPPAktif = LastMonthCloseNPPAktif,
        @LastMonthCloseNPPOD = LastMonthCloseNPPOD,
        @LastMonthClosePersenNPPOD = LastMonthClosePersenNPPOD
        from #ODParentBranch
	    where Jaringanname in (select Jaringanname from #temp where BranchID = @branchid )

SET @MsgBody5 = @MsgBody5 + 
					'
						<tr>
						<td class="title" align=center> <b>' +@parentbranch+ '</b> </td>		
						<td class="td" align=center> ' +iif(@OpenNPPAktif=0,'0',convert(varchar(18),format(isnull(@OpenNPPAktif,0),'###,###')))+ '</td>		
						<td class="td" align=center> ' +iif(@OpenNPPOD=0,'0',convert(varchar(18),format(isnull(@OpenNPPOD,0),'###,###')))+ '</td>		
						<td class="important" align=center> <b>' +iif(@OpenPersenNPPOD=0,'0',convert(varchar(18),isnull(@OpenPersenNPPOD,0)))+ ' %</b> </td>	
						<td class="td" align=center> ' +iif(@CloseNPPAktif=0,'0',convert(varchar(18),format(isnull(@CloseNPPAktif,0),'###,###')))+ '</td>		
						<td class="td" align=center> ' +iif(@CloseNPPOD=0,'0',convert(varchar(18),format(isnull(@CloseNPPOD,0),'###,###')))+ '</td>		
						<td class="important" align=center> <b>' +iif(@ClosePersenNPPOD=0,'0',convert(varchar(18),isnull(@ClosePersenNPPOD,0)))+ ' %</b> </td>	
						<td class="important" align=center> <b>' +iif(@ClosePersenNPPOD-@OpenPersenNPPOD=0,'0',convert(varchar(18),isnull(@ClosePersenNPPOD-@OpenPersenNPPOD,0)))+ ' %</b> </td>	
						<td class="td" align=center>' +iif(@LastMonthOpenNPPAktif=0,'0',convert(varchar(18),format(isnull(@LastMonthOpenNPPAktif,0),'###,###')))+ '</td>		
						<td class="td" align=center>' +iif(@LastMonthOpenNPPOD=0,'0',convert(varchar(18),format(isnull(@LastMonthOpenNPPOD,0),'###,###')))+ '</td>		
						<td class="important" align=center> <b>' +iif(@LastMonthOpenPersenNPPOD=0,'0',convert(varchar(18),isnull(@LastMonthOpenPersenNPPOD,0)))+ ' %</b> </td>	
						<td class="td" align=center> ' +iif(@LastMonthCloseNPPAktif=0,'0',convert(varchar(18),format(isnull(@LastMonthCloseNPPAktif,0),'###,###')))+ '</td>		
						<td class="td" align=center> ' +iif(@LastMonthCloseNPPOD=0,'0',convert(varchar(18),format(isnull(@LastMonthCloseNPPOD,0),'###,###')))+ '</td>		
						<td class="important" align=center> <b>' +iif(@LastMonthClosePersenNPPOD=0,'0',convert(varchar(18),isnull(@LastMonthClosePersenNPPOD,0)))+ ' %</b> </td>	
						<td class="important" align=center> <b>' +iif(@LastMonthClosePersenNPPOD-@LastMonthOpenPersenNPPOD=0,'0',convert(varchar(18),isnull(@LastMonthClosePersenNPPOD-@LastMonthOpenPersenNPPOD,0)))+ ' %</b> </td>
						</tr>
					'

			
                
                SELECT @TotalOpenAktifparentbranch = ISNULL(SUM(OpenNPPAktif), 0) FROM #ODParentBranch  where Jaringanname in (select Jaringanname from #temp where BranchID = @branchid )
                SELECT @TotalOpenODparentbranch = ISNULL(SUM(OpenNPPOd), 0) FROM #ODParentBranch  where Jaringanname in (select Jaringanname from #temp where BranchID = @branchid )
                SET @TotalOpenPersenODparentbranch = ISNULL(ROUND(CAST(@TotalOpenODparentbranch AS FLOAT) / NULLIF(CAST(@TotalOpenAktifparentbranch AS FLOAT), 0) * 100, 2), 0)
                
                SELECT @TotalCloseAktifparentbranch = ISNULL(SUM(CloseNPPAktif), 0) FROM #ODParentBranch  where Jaringanname in (select Jaringanname from #temp where BranchID = @branchid )
                SELECT @TotalCloseODparentbranch = ISNULL(SUM(CloseNPPOd), 0) FROM #ODParentBranch  where Jaringanname in (select Jaringanname from #temp where BranchID = @branchid )
                SET @TotalClosePersenODparentbranch = ISNULL(ROUND(CAST(@TotalCloseODparentbranch AS FLOAT) /  NULLIF(CAST(@TotalCloseAktifparentbranch AS FLOAT), 0) * 100, 2), 0)
                
                SELECT @TotalLastMonthOpenAktifparentbranch = ISNULL(SUM(LastMonthOpenNPPAktif), 0) FROM #ODParentBranch  where Jaringanname in (select Jaringanname from #temp where BranchID = @branchid )
                SELECT @TotalLastMonthOpenODparentbranch = ISNULL(SUM(LastMonthOpenNPPOd), 0) FROM #ODParentBranch  where Jaringanname in (select Jaringanname from #temp where BranchID = @branchid )
                SET @TotalLastMonthOpenPersenODparentbranch = ISNULL(ROUND(CAST(@TotalLastMonthOpenODparentbranch AS FLOAT) / NULLIF(CAST(@TotalLastMonthOpenAktifparentbranch AS FLOAT), 0) * 100, 2), 0)
                
                SELECT @TotalLastMonthCloseAktifparentbranch = ISNULL(SUM(LastMonthCloseNPPAktif), 0) FROM #ODParentBranch  where Jaringanname in (select Jaringanname from #temp where BranchID = @branchid )
                SELECT @TotalLastMonthCloseODparentbranch = ISNULL(SUM(LastMonthCloseNPPOd), 0) FROM #ODParentBranch  where Jaringanname in (select Jaringanname from #temp where BranchID = @branchid )
                SET @TotalLastMonthClosePersenODparentbranch = ISNULL(ROUND(CAST(@TotalLastMonthCloseODparentbranch AS FLOAT) / NULLIF(CAST(@TotalLastMonthCloseAktifparentbranch AS FLOAT), 0) * 100, 2), 0)

				SET @MsgBody5 = @MsgBody5 + 
						'
							<tr>
							<td class="head" align=center> <b> Total </b> </td>	
							<td class="head" align=center> ' +iif(@TotalOpenAktifparentbranch=0,'0',convert(varchar(18),format(isnull(@TotalOpenAktifparentbranch,0),'###,###')))+ '</td>		
							<td class="head" align=center> ' +iif(@TotalOpenODparentbranch=0,'0',convert(varchar(18),format(isnull(@TotalOpenODparentbranch,0),'###,###')))+ '</td>		
							<td class="head" align=center> <b>' +iif(@TotalOpenPersenODparentbranch=0,'0',convert(varchar(18),isnull(@TotalOpenPersenODparentbranch,0)))+ ' %</b> </td>	
							<td class="head" align=center> ' +iif(@TotalCloseAktifparentbranch=0,'0',convert(varchar(18),format(isnull(@TotalCloseAktifparentbranch,0),'###,###')))+ '</td>		
							<td class="head" align=center> ' +iif(@TotalCloseODparentbranch=0,'0',convert(varchar(18),format(isnull(@TotalCloseODparentbranch,0),'###,###')))+ '</td>		
							<td class="head" align=center> <b>' +iif(@TotalClosePersenODparentbranch=0,'0',convert(varchar(18),isnull(@TotalClosePersenODparentbranch,0)))+ ' %</b> </td>	
							<td class="head" align=center> <b>' +iif(@TotalClosePersenODparentbranch-@TotalOpenPersenODparentbranch=0,'0',convert(varchar(18),isnull(@TotalClosePersenODparentbranch-@TotalOpenPersenODparentbranch,0)))+ ' %</b> </td>
							<td class="head" align=center> ' +iif(@TotalLastMonthOpenAktifparentbranch=0,'0',convert(varchar(18),format(isnull(@TotalLastMonthOpenAktifparentbranch,0),'###,###')))+ '</td>		
							<td class="head" align=center> ' +iif(@TotalLastMonthOpenODparentbranch=0,'0',convert(varchar(18),format(isnull(@TotalLastMonthOpenODparentbranch,0),'###,###')))+ '</td>		
							<td class="head" align=center> <b>' +iif(@TotalLastMonthOpenPersenODparentbranch=0,'0',convert(varchar(18),isnull(@TotalLastMonthOpenPersenODparentbranch,0)))+ ' %</b> </td>	
							<td class="head" align=center> ' +iif(@TotalLastMonthCloseAktifparentbranch=0,'0',convert(varchar(18),format(isnull(@TotalLastMonthCloseAktifparentbranch,0),'###,###')))+ '</td>		
							<td class="head" align=center> ' +iif(@TotalLastMonthCloseODparentbranch=0,'0',convert(varchar(18),format(isnull(@TotalLastMonthCloseODparentbranch,0),'###,###')))+ '</td>		
							<td class="head" align=center> <b>' +iif(@TotalLastMonthClosePersenODparentbranch=0,'0',convert(varchar(18),isnull(@TotalLastMonthClosePersenODparentbranch,0)))+ ' %</b> </td>	
							<td class="head" align=center> <b>' +iif(@TotalLastMonthClosePersenODparentbranch-@TotalLastMonthOpenPersenODparentbranch=0,'0',convert(varchar(18),isnull(@TotalLastMonthClosePersenODparentbranch-@TotalLastMonthOpenPersenODparentbranch,0)))+ ' %</b> </td>
							</tr>
						'
			
set @MsgBody6 = '' 
set @MsgHdr6 = '  </style>    
		           
	            <!-- TABLE Cabang pos -->
					</table>
					<br>
                    <table ID="Cabang Pos" class="table">
					     <tr>
						   <td colspan="15" class="title" align=center> <b> Unit Over Due By Cabang Pos </b> </td>
						</tr>
						<tr>
							<td rowspan="3" class="head"> <b> Cabang Pos </b> </td>   
							<td colspan="7" class="head"> <b>' + convert(varchar(20),convert(datetime,@mtdThisMonth),106) + '</b> </td>
							<td colspan="7" class="head"> <b>' + convert(varchar(20),convert(datetime,@mtdlastMonth),106) + '</b> </td>							
						</tr>
						<tr>
							<td colspan="3" class="head"> <b> Open </b> </td>
							<td colspan="3" class="head"> <b> Close </b> </td>
							<td rowspan="2" class="head"> <b> Changes </b> </td>
							<td colspan="3" class="head"> <b> Open </b> </td>
							<td colspan="3" class="head"> <b> Close </b> </td>
							<td rowspan="2" class="head"> <b> Changes </b> </td>

						</tr>
						<tr>
							<td class="head"> <b> Unit Active </b> </td>
							<td class="head"> <b> Unit OD     </b> </td>  
							<td class="head"> <b> % OD        </b> </td>    
							<td class="head"> <b> Unit Active </b> </td>
							<td class="head"> <b> Unit OD     </b> </td>  
							<td class="head"> <b> % OD        </b> </td>  		
							<td class="head"> <b> Unit Active </b> </td>
							<td class="head"> <b> Unit OD     </b> </td>  
							<td class="head"> <b> % OD        </b> </td>    
							<td class="head"> <b> Unit Active </b> </td>
							<td class="head"> <b> Unit OD     </b> </td>  
							<td class="head"> <b> % OD        </b> </td>  											
						</tr>
            '

DECLARE cur2 CURSOR FOR 
SELECT * FROM #ODcabangpos where Branchname in (select Branchname from #temp where branchid = @branchid ) ;

OPEN cur2;

FETCH NEXT FROM cur2 INTO  @branchname, @OpenNPPaktif,@OpenNPPod,@OpenPersenNPPOD,@CloseNPPaktif,@CloseNPPod,@ClosePersenNPPOD,@LastMonthOpenNPPaktif,
@LastMonthOpenNPPod,@LastMonthOpenPersenNPPOD,@LastMonthCloseNPPaktif,@LastMonthCloseNPPod,@LastMonthClosePersenNPPOD 
WHILE @@FETCH_STATUS = 0																		   
BEGIN

	

SET @MsgBody6 = @MsgBody6 + 
					'
						<tr>
						<td class="title" align=center> <b>' +@BranchName+ '</b> </td>		
						<td class="td" align=center> ' +iif(@OpenNPPAktif=0,'0',convert(varchar(18),format(isnull(@OpenNPPAktif,0),'###,###')))+ '</td>		
						<td class="td" align=center> ' +iif(@OpenNPPOD=0,'0',convert(varchar(18),format(isnull(@OpenNPPOD,0),'###,###')))+ '</td>		
						<td class="important" align=center> <b>' +iif(@OpenPersenNPPOD=0,'0',convert(varchar(18),isnull(@OpenPersenNPPOD,0)))+ ' %</b> </td>	
						<td class="td" align=center> ' +iif(@CloseNPPAktif=0,'0',convert(varchar(18),format(isnull(@CloseNPPAktif,0),'###,###')))+ '</td>		
						<td class="td" align=center> ' +iif(@CloseNPPOD=0,'0',convert(varchar(18),format(isnull(@CloseNPPOD,0),'###,###')))+ '</td>		
						<td class="important" align=center> <b>' +iif(@ClosePersenNPPOD=0,'0',convert(varchar(18),isnull(@ClosePersenNPPOD,0)))+ ' %</b> </td>	
						<td class="important" align=center> <b>' +iif(@ClosePersenNPPOD-@OpenPersenNPPOD=0,'0',convert(varchar(18),isnull(@ClosePersenNPPOD-@OpenPersenNPPOD,0)))+ ' %</b> </td>	
						<td class="td" align=center>' +iif(@LastMonthOpenNPPAktif=0,'0',convert(varchar(18),format(isnull(@LastMonthOpenNPPAktif,0),'###,###')))+ '</td>		
						<td class="td" align=center>' +iif(@LastMonthOpenNPPOD=0,'0',convert(varchar(18),format(isnull(@LastMonthOpenNPPOD,0),'###,###')))+ '</td>		
						<td class="important" align=center> <b>' +iif(@LastMonthOpenPersenNPPOD=0,'0',convert(varchar(18),isnull(@LastMonthOpenPersenNPPOD,0)))+ ' %</b> </td>	
						<td class="td" align=center> ' +iif(@LastMonthCloseNPPAktif=0,'0',convert(varchar(18),format(isnull(@LastMonthCloseNPPAktif,0),'###,###')))+ '</td>		
						<td class="td" align=center> ' +iif(@LastMonthCloseNPPOD=0,'0',convert(varchar(18),format(isnull(@LastMonthCloseNPPOD,0),'###,###')))+ '</td>		
						<td class="important" align=center> <b>' +iif(@LastMonthClosePersenNPPOD=0,'0',convert(varchar(18),isnull(@LastMonthClosePersenNPPOD,0)))+ ' %</b> </td>	
						<td class="important" align=center> <b>' +iif(@LastMonthClosePersenNPPOD-@LastMonthOpenPersenNPPOD=0,'0',convert(varchar(18),isnull(@LastMonthClosePersenNPPOD-@LastMonthOpenPersenNPPOD,0)))+ ' %</b> </td>
						</tr>
					'
FETCH NEXT FROM cur2 INTO @branchname, @OpenNPPaktif,@OpenNPPod,@OpenPersenNPPOD,@CloseNPPaktif,@CloseNPPod,@ClosePersenNPPOD,@LastMonthOpenNPPaktif,
@LastMonthOpenNPPod,@LastMonthOpenPersenNPPOD,@LastMonthCloseNPPaktif,@LastMonthCloseNPPod,@LastMonthClosePersenNPPOD 
END
				
                
                select @TotalOpenAktifBranchName=sum(OpenNPPAktif) from #ODcabangpos where Branchname in (select Branchname from #temp where branchid = @branchid )
				select @TotalOpenODBranchName=sum(OpenNPPOd) from #ODcabangpos where Branchname in (select Branchname from #temp where branchid = @branchid )
				set @TotalOpenPersenODBranchName = round(CAST(@TotalOpenODBranchName as float)/CAST(@TotalOpenAktifBranchName as float)*100,2)

				select @TotalCloseAktifBranchName=sum(CloseNPPAktif) from #ODcabangpos where Branchname in (select Branchname from #temp where branchid = @branchid)
				select @TotalCloseODBranchName=sum(CloseNPPOd) from #ODcabangpos where Branchname in (select Branchname from #temp where branchid = @branchid )
				set @TotalClosePersenODBranchName = round(CAST(@TotalCloseODBranchName as float)/CAST(@TotalCloseAktifBranchName as float)*100,2)
			
				select @TotalLastMonthOpenAktifBranchName=sum(LastMonthOpenNPPAktif) from #ODcabangpos where Branchname in (select Branchname from #temp where branchid = @branchid)
				select @TotalLastMonthOpenODBranchName=sum(LastMonthOpenNPPOd) from #ODcabangpos where Branchname in (select Branchname from #temp where branchid = @branchid)
				set @TotalLastMonthOpenPersenODBranchName = round(CAST(@TotalLastMonthOpenODBranchName as float)/CAST(@TotalLastMonthOpenAktifBranchName as float)*100,2)

				select @TotalLastMonthCloseAktifBranchName=sum(LastMonthCloseNPPAktif) from #ODcabangpos where Branchname in (select Branchname from #temp where branchid = @branchid)
				select @TotalLastMonthCloseODBranchName=sum(LastMonthCloseNPPOd) from #ODcabangpos where Branchname in (select Branchname from #temp where branchid = @branchid)
				set @TotalLastMonthClosePersenODBranchName = round(CAST(@TotalLastMonthCloseODBranchName as float)/CAST(@TotalLastMonthCloseAktifBranchName as float)*100,2)

				SET @MsgBody6 = @MsgBody6 + 
						'
							<tr>
							<td class="head" align=center> <b> Total </b> </td>	
							<td class="head" align=center> ' +iif(@TotalOpenAktifBranchName=0,'0',convert(varchar(18),format(isnull(@TotalOpenAktifBranchName,0),'###,###')))+ '</td>		
							<td class="head" align=center> ' +iif(@TotalOpenODBranchName=0,'0',convert(varchar(18),format(isnull(@TotalOpenODBranchName,0),'###,###')))+ '</td>		
							<td class="head" align=center> <b>' +iif(@TotalOpenPersenODBranchName=0,'0',convert(varchar(18),isnull(@TotalOpenPersenODBranchName,0)))+ ' %</b> </td>	
							<td class="head" align=center> ' +iif(@TotalCloseAktifBranchName=0,'0',convert(varchar(18),format(isnull(@TotalCloseAktifBranchName,0),'###,###')))+ '</td>		
							<td class="head" align=center> ' +iif(@TotalCloseODBranchName=0,'0',convert(varchar(18),format(isnull(@TotalCloseODBranchName,0),'###,###')))+ '</td>		
							<td class="head" align=center> <b>' +iif(@TotalClosePersenODBranchName=0,'0',convert(varchar(18),isnull(@TotalClosePersenODBranchName,0)))+ ' %</b> </td>	
							<td class="head" align=center> <b>' +iif(@TotalClosePersenODBranchName-@TotalOpenPersenODBranchName=0,'0',convert(varchar(18),isnull(@TotalClosePersenODBranchName-@TotalOpenPersenODBranchName,0)))+ ' %</b> </td>
							<td class="head" align=center> ' +iif(@TotalLastMonthOpenAktifBranchName=0,'0',convert(varchar(18),format(isnull(@TotalLastMonthOpenAktifBranchName,0),'###,###')))+ '</td>		
							<td class="head" align=center> ' +iif(@TotalLastMonthOpenODBranchName=0,'0',convert(varchar(18),format(isnull(@TotalLastMonthOpenODBranchName,0),'###,###')))+ '</td>		
							<td class="head" align=center> <b>' +iif(@TotalLastMonthOpenPersenODBranchName=0,'0',convert(varchar(18),isnull(@TotalLastMonthOpenPersenODBranchName,0)))+ ' %</b> </td>	
							<td class="head" align=center> ' +iif(@TotalLastMonthCloseAktifBranchName=0,'0',convert(varchar(18),format(isnull(@TotalLastMonthCloseAktifBranchName,0),'###,###')))+ '</td>		
							<td class="head" align=center> ' +iif(@TotalLastMonthCloseODBranchName=0,'0',convert(varchar(18),format(isnull(@TotalLastMonthCloseODBranchName,0),'###,###')))+ '</td>		
							<td class="head" align=center> <b>' +iif(@TotalLastMonthClosePersenODBranchName=0,'0',convert(varchar(18),isnull(@TotalLastMonthClosePersenODBranchName,0)))+ ' %</b> </td>	
							<td class="head" align=center> <b>' +iif(@TotalLastMonthClosePersenODBranchName-@TotalLastMonthOpenPersenODBranchName=0,'0',convert(varchar(18),isnull(@TotalLastMonthClosePersenODBranchName-@TotalLastMonthOpenPersenODBranchName,0)))+ ' %</b> </td>
							</tr>
						'



close cur2
deallocate cur2

---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


	set @MsgFtr = ' 
		                </table>
						<br>

						<!-- TABLE INFO GENERATE -->
						<table ID="InfoGenerate" class="table">					
						<tr>
                             <td colspan="7" class="head"> <i> Generated on :   ' + convert(varchar(12),getdate(),106) + space(4) + convert(varchar(12),getdate(),108) + space(10) +' </i></td>
                        </tr>
						</table>
					   </body>
                      </html>
                      '

SET @mailbody = @MsgHdr1 + @MsgBody1+@Pemisah+ @MsgHdr2 + @MsgBody2 + @Pemisah+ @MsgHdr3 + @MsgBody3 +@Pemisah+ @MsgHdr4 + @MsgBody4 +@Pemisah+ @MsgHdr5 + @MsgBody5 + @Pemisah+ @MsgHdr6 + @MsgBody6 +@Pemisah+ @MsgFtr



declare @starting int = 1, @max int

select @starting = 1, @max = (select len(@mailbody))

while @starting<= @max
begin
print (substring(@mailbody,@starting,8000))
set @starting +=8000
end

--select * from #tempPT
--select * from #wilayah
--select * from #regional
--select * from #area
--select * from #ODParentBranch
--select * from #ODcabangpos

--456
--select areaname from #m
--where branchid = '456' 




--select * from #m 
--where areaname in (
--	select areaname from #m 
--	where branchid = '456' )


--select * from #area


--select * from #m 
--where wilayahname in (
--	select wilayahname from #m 
--	where branchid = '456' )



--select * from #m 
--where branchid = '456'

--select * from #m 
--where hbaname in (
--	select hbaname from #m 
--	where branchid = '456' )







