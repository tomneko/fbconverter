{*******************************************************}
{                                                       }
{              Firebird Database Converter              }
{                                                       }
{*******************************************************}

// -----------------------------------------------------------------------------
// Note:
//   - String Constant Unit
// -----------------------------------------------------------------------------
unit uFBConvConsts;

interface

const
  // for Error Dialog
  ERR_MSG_FILENOTEXISTS = 'Source database file not exits.';
  ERR_MSG_DIRECTORYNOTEXISTS = 'Destination database directory not exits.';
  ERR_MSG_DATABASEOPENERROR = 'Can''t open file.';
  ERR_MSG_INTERNALCHARSETNOTSELECT = 'Internal charset is not selected.';
  ERR_MSG_SAMEFILE = 'Same file has been specified.';
  ERR_MSG_CANTCREATEDATABASE = 'Can''t Create Database';
  ERR_MSG_REMOTEPATHSPECIFIED = 'Can not specified remote path.';

  // for Information Dialog
  INF_MSG_DONE = 'Done.';

  // for Confirmation Dialog
  CON_MSG_FILEEXIST = 'Destination database file already exist.' + #$0D#$0A + 'Overwrite?';
  CON_MSG_VERSIONMISMATCH = '%s Client/Server version mismatch.' + #$0D#$0A + // Source / Destination
                            '  Client: %s' + #$0D#$0A +                         // Client Version
                            '  Server: %s' + #$0D#$0A;                          // Server Version
  CON_MSG_VERSIONMISMATCH1 = CON_MSG_VERSIONMISMATCH + 'Do you want to continue?';
  CON_MSG_VERSIONMISMATCH2 = CON_MSG_VERSIONMISMATCH + 'Fix it?';

  // Misc
  SOURCE_STR = 'Source';
  DESTINATON_STR = 'Destination';
  MSG_STARTCONVERSION = 'Start conversion.';
  MSG_ENDCONVERSION = 'Done.';
  MSG_PROCESSED_RECORD = '  %s records processed. %s';                          // Records / Process time
  MSG_PROCESSED_RECORD_BLOCK = '%s-%s record block processed. %s';              // Start record / End record / Process time
  MSG_UNKNOWN_DATATYPE = 'Unknown data type.';
  MSG_SERVER_VERSION = 'Server Version: %s';                                    // Server Version
  MSG_CLIENT_VERSION = 'Client Version: %s';                                    // Client Version
  MSG_ODS_VERSION = 'On Disk Structure: %s';                                    // ODS Version
  MSG_PAGE_SIZE = 'PageSize: %d';                                               // Page Size
  MSG_DIALECT = 'Dialect: %d';                                                  // Dialect
  MSG_CONNECTON_SUCCESS = 'Connection success.';
  MSG_CONNECTON_SUCCESS2 = MSG_CONNECTON_SUCCESS + #$0D#$0A + 'But,';

implementation

end.
