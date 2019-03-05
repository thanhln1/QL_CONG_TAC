USE [QL_CONG_TAC_PHI]
GO

/****** Object:  Table [dbo].[MT_CONG_TY]    Script Date: 3/5/2019 4:27:50 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[MT_CONG_TY](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[MA_KHACH_HANG] [nvarchar](50) NOT NULL,
	[KHACH_HANG] [nvarchar](50) NOT NULL,
	[MAU] [nvarchar](50) NULL,
	[THANG] [nvarchar](50) NULL
) ON [PRIMARY]

GO


