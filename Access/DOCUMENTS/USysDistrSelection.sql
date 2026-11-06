CREATE TABLE USysDistrSelection (
    SessionID TEXT(36) NOT NULL,
    RecipientID LONG NOT NULL,
    Copies MEMO,
    CreatedAt DATETIME NOT NULL,
    CONSTRAINT PK_USysDistrSelection PRIMARY KEY (SessionID, RecipientID)
);
