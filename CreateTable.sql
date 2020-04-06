CREATE TABLE Main (
	ID int NOT NULL PRIMARY KEY IDENTITY(1,1),
	[TBM Number] nvarchar(255) NOT NULL,
	[Defect Code] nvarchar(255) NOT NULL,
	[Fault Description] nvarchar(255),
	[Idle Time] int NOT NULL,
	Operator nvarchar(255) NOT NULL,
	[Description Notes] nvarchar(MAX),
	Date date NOT NULL,
	[Shift Number] int NOT NULL,
	[Shift Time] nvarchar(255) NOT NULL
);