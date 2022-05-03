USE [MetalBase]
GO

/****** Object:  Table [dbo].[Marks]    Script Date: 12.10.2019 10:16:12 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[MarksNull](
	[MarkID] [int] IDENTITY(1,1) NOT NULL,
	[MarkType] [varchar] (100),
	[Mark] [varchar](500))
GO
Insert into MarksNull (Mark)
select Mark from MarksNerj
GO
Update MarksNull
set MarkType = 'Nerj'
GO
CREATE TABLE [dbo].[Marks](
	[MarkID] [int] IDENTITY(1,1) NOT NULL,
	[MarkType] [varchar](100) NOT NULL,
	[Mark] [varchar](500) NOT NULL
) ON [PRIMARY]
GO
insert into Marks (MarkType, Mark)
select MarkType, Mark from MarksNull
GO
Drop table MarksNull
GO

CREATE PROCEDURE insMarks 
	-- Add the parameters for the stored procedure here
	@Type varchar(100) = 'NotKnown', 
	@Mark varchar(500) = 'ст.12345'
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
	if not exists (select [Mark] from dbo.Marks where Mark = @Mark)
		insert into dbo.Marks (MarkType, Mark)
		values (@Type, @Mark)
	
END
GO
