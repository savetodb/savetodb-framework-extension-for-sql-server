-- =============================================
-- SaveToDB Framework Extension for SQL Server
-- Version 10.8, January 9, 2023
--
-- This script updates SaveToDB Framework Extension 10 to the latest version
--
-- Copyright 2022-2023 Gartle LLC
--
-- License: MIT
-- =============================================

IF 1008 <= COALESCE((SELECT CAST(LEFT(HANDLER_CODE, CHARINDEX('.', HANDLER_CODE) - 1) AS int) * 100 + CAST(RIGHT(HANDLER_CODE, LEN(HANDLER_CODE) - CHARINDEX('.', HANDLER_CODE)) AS float) FROM xls.handlers WHERE TABLE_SCHEMA = 'xls' AND TABLE_NAME = 'savetodb_framework_extension' AND COLUMN_NAME = 'version' AND EVENT_NAME = 'Information'), 0)
    RAISERROR('SaveToDB Framework is up-to-date. Update skipped', 11, 0)
GO

DELETE FROM xls.handlers WHERE TABLE_SCHEMA = 'xls' AND EVENT_NAME = 'Actions' AND HANDLER_NAME IN ('xl_actions_set_framework_10_mode', 'xl_actions_set_framework_9_mode');
GO

UPDATE xls.handlers
SET
    HANDLER_TYPE = s.HANDLER_TYPE
    , HANDLER_CODE = s.HANDLER_CODE
    , TARGET_WORKSHEET = s.TARGET_WORKSHEET
    , MENU_ORDER = s.MENU_ORDER
    , EDIT_PARAMETERS = s.EDIT_PARAMETERS
FROM
    (
    SELECT
        CAST(NULL AS nvarchar) AS TABLE_SCHEMA
        , CAST(NULL AS nvarchar) AS TABLE_NAME
        , CAST(NULL AS nvarchar) AS COLUMN_NAME
        , CAST(NULL AS nvarchar) AS EVENT_NAME
        , CAST(NULL AS nvarchar) AS HANDLER_SCHEMA
        , CAST(NULL AS nvarchar) AS HANDLER_NAME
        , CAST(NULL AS nvarchar) AS HANDLER_TYPE
        , CAST(NULL AS nvarchar) HANDLER_CODE
        , CAST(NULL AS nvarchar) AS TARGET_WORKSHEET
        , CAST(NULL AS int) AS MENU_ORDER
        , CAST(NULL AS bit) AS EDIT_PARAMETERS

    UNION ALL SELECT N'xls', N'savetodb_framework_extension', N'version', N'Information', NULL, NULL, N'ATTRIBUTE', N'10.8', NULL, NULL, NULL
    UNION ALL SELECT N'xls', N'users', NULL, N'Actions', N'xls', N'xl_actions_set_extended_role_permissions', N'PROCEDURE', NULL, N'_MsgBox', 22, 1
    UNION ALL SELECT N'xls', N'users', NULL, N'Actions', N'xls', N'xl_actions_revoke_extended_role_permissions', N'PROCEDURE', NULL, N'_MsgBox', 23, 1
    UNION ALL SELECT N'xls', N'view_columns', NULL, N'Actions', N'xls', N'Developer Guide', N'HTTP', N'https://www.savetodb.com/dev-guide/xls-columns.htm', NULL, 13, NULL
    UNION ALL SELECT N'xls', N'view_formats', NULL, N'Actions', N'xls', N'Developer Guide', N'HTTP', N'https://www.savetodb.com/dev-guide/xls-formats.htm', NULL, 13, NULL
    UNION ALL SELECT N'xls', N'view_handlers', NULL, N'Actions', N'xls', N'Developer Guide', N'HTTP', N'https://www.savetodb.com/dev-guide/xls-handlers.htm', NULL, 13, NULL
    UNION ALL SELECT N'xls', N'view_objects', NULL, N'Actions', N'xls', N'Developer Guide', N'HTTP', N'https://www.savetodb.com/dev-guide/xls-objects.htm', NULL, 13, NULL
    UNION ALL SELECT N'xls', N'view_queries', NULL, N'Actions', N'xls', N'Developer Guide', N'HTTP', N'https://www.savetodb.com/dev-guide/xls-queries.htm', NULL, 13, NULL
    UNION ALL SELECT N'xls', N'view_translations', NULL, N'Actions', N'xls', N'Developer Guide', N'HTTP', N'https://www.savetodb.com/dev-guide/xls-translations.htm', NULL, 13, NULL
    UNION ALL SELECT N'xls', N'view_workbooks', NULL, N'Actions', N'xls', N'Developer Guide', N'HTTP', N'https://www.savetodb.com/dev-guide/xls-workbooks.htm', NULL, 13, NULL

    ) s
    LEFT OUTER JOIN xls.handlers t ON
        t.TABLE_SCHEMA = s.TABLE_SCHEMA
        AND t.TABLE_NAME = s.TABLE_NAME
        AND COALESCE(t.COLUMN_NAME, '') = COALESCE(s.COLUMN_NAME, '')
        AND t.EVENT_NAME = s.EVENT_NAME
        AND COALESCE(t.HANDLER_SCHEMA, '') = COALESCE(s.HANDLER_SCHEMA, '')
        AND COALESCE(t.HANDLER_NAME, '') = COALESCE(s.HANDLER_NAME, '')
WHERE
    NOT COALESCE(t.HANDLER_TYPE, '') = COALESCE(s.HANDLER_TYPE, '')
    OR NOT COALESCE(t.HANDLER_CODE, '')  = COALESCE(s.HANDLER_CODE, '')
    OR NOT COALESCE(t.TARGET_WORKSHEET, '') = COALESCE(s.TARGET_WORKSHEET, '')
    OR NOT COALESCE(t.MENU_ORDER, -1) = COALESCE(s.MENU_ORDER, -1)
    OR NOT COALESCE(t.EDIT_PARAMETERS, 0) = COALESCE(s.EDIT_PARAMETERS, 0);
GO

INSERT INTO xls.handlers (TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, EVENT_NAME, HANDLER_SCHEMA, HANDLER_NAME, HANDLER_TYPE, HANDLER_CODE, TARGET_WORKSHEET, MENU_ORDER, EDIT_PARAMETERS)
SELECT s.TABLE_SCHEMA, s.TABLE_NAME, s.COLUMN_NAME, s.EVENT_NAME, s.HANDLER_SCHEMA, s.HANDLER_NAME, s.HANDLER_TYPE, s.HANDLER_CODE, s.TARGET_WORKSHEET, s.MENU_ORDER, s.EDIT_PARAMETERS
FROM
    (
    SELECT
        CAST(NULL AS nvarchar) AS TABLE_SCHEMA
        , CAST(NULL AS nvarchar) AS TABLE_NAME
        , CAST(NULL AS nvarchar) AS COLUMN_NAME
        , CAST(NULL AS nvarchar) AS EVENT_NAME
        , CAST(NULL AS nvarchar) AS HANDLER_SCHEMA
        , CAST(NULL AS nvarchar) AS HANDLER_NAME
        , CAST(NULL AS nvarchar) AS HANDLER_TYPE
        , CAST(NULL AS nvarchar) HANDLER_CODE
        , CAST(NULL AS nvarchar) AS TARGET_WORKSHEET
        , CAST(NULL AS int) AS MENU_ORDER
        , CAST(NULL AS bit) AS EDIT_PARAMETERS

    UNION ALL SELECT N'xls', N'savetodb_framework_extension', N'version', N'Information', NULL, NULL, N'ATTRIBUTE', N'10.8', NULL, NULL, NULL
    UNION ALL SELECT N'xls', N'users', NULL, N'Actions', N'xls', N'xl_actions_set_extended_role_permissions', N'PROCEDURE', NULL, N'_MsgBox', 22, 1
    UNION ALL SELECT N'xls', N'users', NULL, N'Actions', N'xls', N'xl_actions_revoke_extended_role_permissions', N'PROCEDURE', NULL, N'_MsgBox', 23, 1
    UNION ALL SELECT N'xls', N'view_columns', NULL, N'Actions', N'xls', N'Developer Guide', N'HTTP', N'https://www.savetodb.com/dev-guide/xls-columns.htm', NULL, 13, NULL
    UNION ALL SELECT N'xls', N'view_formats', NULL, N'Actions', N'xls', N'Developer Guide', N'HTTP', N'https://www.savetodb.com/dev-guide/xls-formats.htm', NULL, 13, NULL
    UNION ALL SELECT N'xls', N'view_handlers', NULL, N'Actions', N'xls', N'Developer Guide', N'HTTP', N'https://www.savetodb.com/dev-guide/xls-handlers.htm', NULL, 13, NULL
    UNION ALL SELECT N'xls', N'view_objects', NULL, N'Actions', N'xls', N'Developer Guide', N'HTTP', N'https://www.savetodb.com/dev-guide/xls-objects.htm', NULL, 13, NULL
    UNION ALL SELECT N'xls', N'view_queries', NULL, N'Actions', N'xls', N'Developer Guide', N'HTTP', N'https://www.savetodb.com/dev-guide/xls-queries.htm', NULL, 13, NULL
    UNION ALL SELECT N'xls', N'view_translations', NULL, N'Actions', N'xls', N'Developer Guide', N'HTTP', N'https://www.savetodb.com/dev-guide/xls-translations.htm', NULL, 13, NULL
    UNION ALL SELECT N'xls', N'view_workbooks', NULL, N'Actions', N'xls', N'Developer Guide', N'HTTP', N'https://www.savetodb.com/dev-guide/xls-workbooks.htm', NULL, 13, NULL

    ) s
    LEFT OUTER JOIN xls.handlers t ON
        t.TABLE_SCHEMA = s.TABLE_SCHEMA
        AND t.TABLE_NAME = s.TABLE_NAME
        AND COALESCE(t.COLUMN_NAME, '') = COALESCE(s.COLUMN_NAME, '')
        AND t.EVENT_NAME = s.EVENT_NAME
        AND COALESCE(t.HANDLER_SCHEMA, '') = COALESCE(s.HANDLER_SCHEMA, '')
        AND COALESCE(t.HANDLER_NAME, '') = COALESCE(s.HANDLER_NAME, '')
        AND COALESCE(t.HANDLER_TYPE, '') = COALESCE(s.HANDLER_TYPE, '')
WHERE
    t.TABLE_NAME IS NULL
    AND s.TABLE_NAME IS NOT NULL;
GO

IF (SELECT ROUTINE_NAME FROM INFORMATION_SCHEMA.ROUTINES WHERE ROUTINE_SCHEMA = 'xls' AND ROUTINE_NAME = 'xl_actions_set_framework_10_mode') IS NULL
EXEC('DROP PROCEDURE [xls].[xl_actions_set_framework_10_mode];')
GO
IF (SELECT ROUTINE_NAME FROM INFORMATION_SCHEMA.ROUTINES WHERE ROUTINE_SCHEMA = 'xls' AND ROUTINE_NAME = 'xl_actions_set_framework_9_mode') IS NULL
EXEC('DROP PROCEDURE [xls].[xl_actions_set_framework_9_mode];')
GO

IF (SELECT ROUTINE_NAME FROM INFORMATION_SCHEMA.ROUTINES WHERE ROUTINE_SCHEMA = 'xls' AND ROUTINE_NAME = 'xl_actions_set_extended_role_permissions') IS NULL
EXEC('CREATE PROCEDURE [xls].[xl_actions_set_extended_role_permissions] AS SET NOCOUNT ON;')
GO
IF (SELECT ROUTINE_NAME FROM INFORMATION_SCHEMA.ROUTINES WHERE ROUTINE_SCHEMA = 'xls' AND ROUTINE_NAME = 'xl_actions_revoke_extended_role_permissions') IS NULL
EXEC('CREATE PROCEDURE [xls].[xl_actions_revoke_extended_role_permissions] AS SET NOCOUNT ON;')
GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     10.4, 2022-10-13
-- Description: Activates using SaveToDB Framework Extension views instead of SaveToDB Framework tables
-- =============================================

ALTER PROCEDURE [xls].[xl_actions_set_extended_role_permissions]
    @confirm bit = 0
AS
BEGIN

SET NOCOUNT ON;

IF COALESCE(@confirm, 0) = 0 RETURN;

GRANT SELECT ON xls.view_columns        TO xls_users;
GRANT SELECT ON xls.view_formats        TO xls_users;
GRANT SELECT ON xls.view_handlers       TO xls_users;
GRANT SELECT ON xls.view_objects        TO xls_users;
GRANT SELECT ON xls.view_translations   TO xls_users;
GRANT SELECT ON xls.view_workbooks      TO xls_users;
GRANT SELECT ON xls.view_queries        TO xls_users;

REVOKE SELECT, VIEW DEFINITION ON xls.columns       FROM xls_users;
REVOKE SELECT, VIEW DEFINITION ON xls.formats       FROM xls_users;
REVOKE SELECT, VIEW DEFINITION ON xls.handlers      FROM xls_users;
REVOKE SELECT, VIEW DEFINITION ON xls.objects       FROM xls_users;
REVOKE SELECT, VIEW DEFINITION ON xls.translations  FROM xls_users;
REVOKE SELECT, VIEW DEFINITION ON xls.workbooks     FROM xls_users;
REVOKE SELECT, VIEW DEFINITION ON xls.queries       FROM xls_users;

GRANT EXECUTE ON xls.xl_update_table_format TO xls_formats;

REVOKE SELECT, INSERT, UPDATE, DELETE, VIEW DEFINITION ON xls.formats FROM xls_formats;

DENY SELECT, INSERT, UPDATE, DELETE, VIEW DEFINITION ON xls.view_columns        TO xls_developers;
DENY SELECT, INSERT, UPDATE, DELETE, VIEW DEFINITION ON xls.view_formats        TO xls_developers;
DENY SELECT, INSERT, UPDATE, DELETE, VIEW DEFINITION ON xls.view_handlers       TO xls_developers;
DENY SELECT, INSERT, UPDATE, DELETE, VIEW DEFINITION ON xls.view_objects        TO xls_developers;
DENY SELECT, INSERT, UPDATE, DELETE, VIEW DEFINITION ON xls.view_translations   TO xls_developers;
DENY SELECT, INSERT, UPDATE, DELETE, VIEW DEFINITION ON xls.view_workbooks      TO xls_developers;
DENY SELECT, INSERT, UPDATE, DELETE, VIEW DEFINITION ON xls.view_queries        TO xls_developers;

END


GO

-- =============================================
-- Author:      Gartle LLC
-- Release:     10.4, 2022-10-13
-- Description: Activates using SaveToDB Framework tables instead of SaveToDB Framework Extension views
-- =============================================

ALTER PROCEDURE [xls].[xl_actions_revoke_extended_role_permissions]
    @confirm bit = 0
AS
BEGIN

SET NOCOUNT ON;

IF COALESCE(@confirm, 0) = 0 RETURN;

REVOKE SELECT ON xls.view_columns        FROM xls_users;
REVOKE SELECT ON xls.view_formats        FROM xls_users;
REVOKE SELECT ON xls.view_handlers       FROM xls_users;
REVOKE SELECT ON xls.view_objects        FROM xls_users;
REVOKE SELECT ON xls.view_translations   FROM xls_users;
REVOKE SELECT ON xls.view_workbooks      FROM xls_users;
REVOKE SELECT ON xls.view_queries        FROM xls_users;

GRANT SELECT, VIEW DEFINITION ON xls.columns        TO xls_users;
GRANT SELECT, VIEW DEFINITION ON xls.formats        TO xls_users;
GRANT SELECT, VIEW DEFINITION ON xls.handlers       TO xls_users;
GRANT SELECT, VIEW DEFINITION ON xls.objects        TO xls_users;
GRANT SELECT, VIEW DEFINITION ON xls.translations   TO xls_users;
GRANT SELECT, VIEW DEFINITION ON xls.workbooks      TO xls_users;
GRANT SELECT, VIEW DEFINITION ON xls.queries        TO xls_users;

REVOKE EXECUTE ON xls.xl_update_table_format FROM xls_formats;

GRANT SELECT, INSERT, UPDATE, DELETE, VIEW DEFINITION ON xls.formats TO xls_formats;

END


GO

print 'SaveToDB Framework Extension updated';
