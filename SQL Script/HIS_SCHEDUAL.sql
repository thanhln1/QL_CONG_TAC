USE [QL_CONG_TAC_PHI]
GO

/****** Object:  Table [dbo].[HIS_SCHEDUAL]    Script Date: 3/5/2019 4:27:03 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[HIS_SCHEDUAL](
	[ID] [int] NOT NULL,
	[MA_NHAN_VIEN] [nvarchar](50) NOT NULL,
	[THANG] [int] NOT NULL,
	[NAM] [int] NOT NULL,
	[TUAN1_THU2] [nvarchar](50) NULL,
	[TUAN1_THU3] [nvarchar](50) NULL,
	[TUAN1_THU4] [nvarchar](50) NULL,
	[TUAN1_THU5] [nvarchar](50) NULL,
	[TUAN1_THU6] [nvarchar](50) NULL,
	[TUAN1_THU7] [nvarchar](50) NULL,
	[TUAN1_CN] [nvarchar](50) NULL,
	[TUAN2_THU2] [nvarchar](50) NULL,
	[TUAN2_THU3] [nvarchar](50) NULL,
	[TUAN2_THU4] [nvarchar](50) NULL,
	[TUAN2_THU5] [nvarchar](50) NULL,
	[TUAN2_THU6] [nvarchar](50) NULL,
	[TUAN2_THU7] [nvarchar](50) NULL,
	[TUAN2_CN] [nvarchar](50) NULL,
	[TUAN3_THU2] [nvarchar](50) NULL,
	[TUAN3_THU3] [nvarchar](50) NULL,
	[TUAN3_THU4] [nvarchar](50) NULL,
	[TUAN3_THU5] [nvarchar](50) NULL,
	[TUAN3_THU6] [nvarchar](50) NULL,
	[TUAN3_THU7] [nvarchar](50) NULL,
	[TUAN3_CN] [nvarchar](50) NULL,
	[TUAN4_THU2] [nvarchar](50) NULL,
	[TUAN4_THU3] [nvarchar](50) NULL,
	[TUAN4_THU4] [nvarchar](50) NULL,
	[TUAN4_THU5] [nvarchar](50) NULL,
	[TUAN4_THU6] [nvarchar](50) NULL,
	[TUAN4_THU7] [nvarchar](50) NULL,
	[TUAN4_CN] [nvarchar](50) NULL
) ON [PRIMARY]

GO


