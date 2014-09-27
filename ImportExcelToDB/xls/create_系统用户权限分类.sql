
/****** Object: Table [dbo].[系统用户权限分类]   Script Date: 2014/9/24 20:51:20 ******/
USE [Gwms_xw_test];
GO
SET ANSI_NULLS ON;
GO
SET QUOTED_IDENTIFIER ON;
GO
CREATE TABLE [dbo].[系统用户权限分类] (
[功能模块] varchar(128) NULL,
[一级功能分类] varchar(128) NULL,
[系统管理员] varchar(128) NULL,
[参数维护员] varchar(128) NULL,
[数据补录员] varchar(128) NULL,
[普通用户] varchar(128) NULL
)
ON [PRIMARY]
WITH (DATA_COMPRESSION = NONE);
GO


                       
