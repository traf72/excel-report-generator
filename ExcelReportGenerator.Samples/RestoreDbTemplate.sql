-- Download database backup from https://github.com/Microsoft/sql-server-samples/releases/download/adventureworks/AdventureWorks2012.bak

-- Restore backup on server (change <BackupPath>, <DbPath> and <LogPath> templates)
RESTORE DATABASE AdventureWorks
FROM DISK = '<BackupPath>\AdventureWorks2012.bak'
WITH MOVE 'AdventureWorks2012' TO '<DbPath>\AdventureWorks2012.mdf',
MOVE 'AdventureWorks2012_log' TO '<LogPath>\AdventureWorks2012_log.ldf';

-- Then change the connection string in App.config if necessary