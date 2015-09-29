{*******************************************************}
{                                                       }
{              Firebird Database Converter              }
{                                                       }
{*******************************************************}

// -----------------------------------------------------------------------------
// Note:
//   - Firebird Raw Access Utility Unit
// -----------------------------------------------------------------------------
unit uFBRawAccess;

interface

uses
  System.Classes, System.SysUtils, ZPlainFirebirdInterbaseConstants, ZPlainLoader;

type
  TISC_SVC_HANDLE = PVoid;
  PISC_SVC_HANDLE = ^TISC_SVC_HANDLE;

  TISC_DB_HANDLE = LongWord;
  PISC_DB_HANDLE = ^TISC_DB_HANDLE;

  ISC_STATUS_VECTOR = array [0..19] of ISC_STATUS;
  PSTATUS_VECTOR    = ^ISC_STATUS_VECTOR;
  PPSTATUS_VECTOR   = ^PSTATUS_VECTOR;

  Tisc_service_attach     = function (status_vector: PSTATUS_VECTOR;
                                      service_length: UShort;
                                      service: PByte;
                                      svc_handle: PISC_SVC_HANDLE;
                                      spb_length: UShort;
                                      spb: PByte): ISC_STATUS; stdcall;
  Tisc_service_detach     = function (status_vector: PSTATUS_VECTOR;
                                      service_handle: PISC_SVC_HANDLE): ISC_STATUS; stdcall;
  Tisc_service_query      = function (status_vector: PSTATUS_VECTOR;
                                      svc_handle: PISC_SVC_HANDLE;
                                      recv_handle: PISC_SVC_HANDLE; // Reserved
                                      send_spb_length: UShort;
                                      send_spb: PByte;
                                      recv_spb_length: UShort;
                                      recv_spb: PByte;
                                      buffer_length: UShort;
                                      buffer: PByte): ISC_STATUS; stdcall;
  Tisc_attach_database    = function (status_vector: PSTATUS_VECTOR;
                                      db_name_length: UShort;
                                      db_name: PByte;
                                      db_handle: PISC_DB_HANDLE;
                                      parm_buffer_length: UShort;
                                      parm_buffer: PByte): ISC_STATUS; stdcall;
  Tisc_detach_database    = function (status_vector: PSTATUS_VECTOR;
                                      db_handle: PISC_DB_HANDLE): ISC_STATUS; stdcall;
  Tisc_database_info      = function (status_vector: PSTATUS_VECTOR;
                                      db_handle: PISC_DB_HANDLE;
                                      item_list_buffer_length: UShort;
                                      item_list_buffer: PByte;
                                      result_buffer_length: UShort;
                                      result_buffer: PByte): ISC_STATUS; stdcall;
  Tisc_interprete         = function (buffer: PByte; status_vector_ptr: PPSTATUS_VECTOR): ISC_STATUS; stdcall;
  Tisc_vax_integer        = function (buffer: PByte; length: Short): ISC_LONG; stdcall;
  Tisc_get_client_version = procedure (buffer: PByte); stdcall;

const
  isc_spb_user_name             = 1;
  isc_spb_sys_user_name         = 2;
  isc_spb_sys_user_name_enc     = 3;
  isc_spb_password              = 4;
  isc_spb_password_enc          = 5;
  isc_spb_command_line          = 6;
  isc_spb_dbname                = 7;
  isc_spb_verbose               = 8;
  isc_spb_options               = 9;
  isc_spb_connect_timeout       = 10;
  isc_spb_dummy_packet_interval = 11;
  isc_spb_sql_role_name         = 12;
  isc_spb_instance_name         = 13;
  isc_spb_user_dbname           = 14;
  isc_spb_auth_dbname           = 15;
  isc_spb_last_spb_constant     = isc_spb_auth_dbname;

  isc_spb_version1                                = 1;
  isc_spb_current_version                         = 2;
  isc_spb_version		                              = isc_spb_current_version;
  isc_spb_user_name_mapped_to_server              = isc_dpb_user_name;
  isc_spb_sys_user_name_mapped_to_server          = isc_dpb_sys_user_name;
  isc_spb_sys_user_name_enc_mapped_to_server      = isc_dpb_sys_user_name_enc;
  isc_spb_password_mapped_to_server               = isc_dpb_password;
  isc_spb_password_enc_mapped_to_server           = isc_dpb_password_enc;
  isc_spb_command_line_mapped_to_server           = 105;
  isc_spb_dbname_mapped_to_server                 = 106;
  isc_spb_verbose_mapped_to_server                = 107;
  isc_spb_options_mapped_to_server                = 108;
  isc_spb_user_dbname_mapped_to_server            = 109;
  isc_spb_auth_dbname_mapped_to_server            = isc_spb_user_dbname_mapped_to_server;
  isc_spb_connect_timeout_mapped_to_server        = isc_dpb_connect_timeout;
  isc_spb_dummy_packet_interval_mapped_to_server  = isc_dpb_dummy_packet_interval;
  isc_spb_sql_role_name_mapped_to_server          = isc_dpb_sql_role_name;
  isc_spb_instance_name_mapped_to_server          = 75;

  isc_info_svc_svr_db_info        = 50;
  isc_info_svc_get_license        = 51;
  isc_info_svc_get_license_mask   = 52;
  isc_info_svc_get_config         = 53;
  isc_info_svc_version            = 54;
  isc_info_svc_server_version     = 55;
  isc_info_svc_implementation     = 56;
  isc_info_svc_capabilities       = 57;
  isc_info_svc_user_dbpath        = 58;
  isc_info_svc_get_env            = 59;
  isc_info_svc_get_env_lock       = 60;
  isc_info_svc_get_env_msg        = 61;
  isc_info_svc_line               = 62;
  isc_info_svc_to_eof             = 63;
  isc_info_svc_timeout            = 64;
  isc_info_svc_get_licensed_users = 65;
  isc_info_svc_limbo_trans        = 66;
  isc_info_svc_running            = 67;
  isc_info_svc_get_users          = 68;
  isc_info_svc_get_db_alias       = 69;

  isc_info_end                     = 1;
  isc_info_truncated               = 2;
  isc_info_error                   = 3;
  isc_info_data_not_ready          = 4;
  isc_info_flag_end                = 127;

  isc_info_db_id                  =  4;
  isc_info_reads                  =  5;
  isc_info_writes                 =  6;
  isc_info_fetches                =  7;
  isc_info_marks                  =  8;
  isc_info_implementation         = 11;
  isc_info_version                = 12;
  isc_info_base_level             = 13;
  isc_info_page_size              = 14;
  isc_info_num_buffers            = 15;
  isc_info_limbo                  = 16;
  isc_info_current_memory         = 17;
  isc_info_max_memory             = 18;
  isc_info_window_turns           = 19;
  isc_info_license                = 20;
  isc_info_allocation             = 21;
  isc_info_attachment_id          = 22;
  isc_info_read_seq_count         = 23;
  isc_info_read_idx_count         = 24;
  isc_info_insert_count           = 25;
  isc_info_update_count           = 26;
  isc_info_delete_count           = 27;
  isc_info_backout_count          = 28;
  isc_info_purge_count            = 29;
  isc_info_expunge_count          = 30;
  isc_info_sweep_interval         = 31;
  isc_info_ods_version            = 32;
  isc_info_ods_minor_version      = 33;
  isc_info_no_reserve             = 34;
  isc_info_logfile                = 35;
  isc_info_cur_logfile_name       = 36;
  isc_info_cur_log_part_offset    = 37;
  isc_info_num_wal_buffers        = 38;
  isc_info_wal_buffer_size        = 39;
  isc_info_wal_ckpt_length        = 40;
  isc_info_wal_cur_ckpt_interval  = 41;
  isc_info_wal_prv_ckpt_fname     = 42;
  isc_info_wal_prv_ckpt_poffset   = 43;
  isc_info_wal_recv_ckpt_fname    = 44;
  isc_info_wal_recv_ckpt_poffset  = 45;
  isc_info_wal_grpc_wait_usecs    = 47;
  isc_info_wal_num_io             = 48;
  isc_info_wal_avg_io_size        = 49;
  isc_info_wal_num_commits        = 50;
  isc_info_wal_avg_grpc_size      = 51;
  isc_info_forced_writes          = 52;
  isc_info_user_names             = 53;
  isc_info_page_errors            = 54;
  isc_info_record_errors          = 55;
  isc_info_bpage_errors           = 56;
  isc_info_dpage_errors           = 57;
  isc_info_ipage_errors           = 58;
  isc_info_ppage_errors           = 59;
  isc_info_tpage_errors           = 60;
  isc_info_set_page_buffers       = 61;
  isc_info_db_SQL_dialect         = 62;
  isc_info_db_read_only           = 63;
  isc_info_db_size_in_pages       = 64;

  procedure Build_PBItem(var PBBuf: TBytesStream; Item: Byte);
  procedure Build_PBString(var PBBuf: TBytesStream; Item: Byte; Contents: string);
  function Build_PB(aUserName, aPassword: string): TBytes;
  function Build_FileSpec(aHostName: string; aPort: uInt16; aDatabaseName: string): string;
  function Build_DatabaseName(aHostName: string; aPort: uInt16; aDatabaseName: string): TBytes;

implementation

procedure Build_PBItem(var PBBuf: TBytesStream; Item: Byte);
// パラメータバッファ用アイテムのビルド
begin
  // Item
  PBBuf.Write(Item, SizeOf(Item));
end;

procedure Build_PBString(var PBBuf: TBytesStream; Item: Byte; Contents: string);
// パラメータバッファ用アイテムと文字列のビルド
var
  B: Byte;
  Buf: TBytes;
begin
  // Item
  Build_PBItem(PBBuf, Item);
  // Contents
  Buf := TEncoding.Convert(TEncoding.Default, TEncoding.ANSI, BytesOf(Contents));
  B := Length(Buf);
  PBBuf.Write(B, SizeOf(B));
  PBBuf.WriteBuffer(Buf, Length(Buf));
end;

function Build_PB(aUserName, aPassword: string): TBytes;
// PB (パラメータバッファ) のビルド
var
  BytesStream: TBytesStream;
begin
  BytesStream := TBytesStream.Create;
  try
    // Parameter
    Build_PBItem(BytesStream, isc_dpb_version1);
    // User Name
    Build_PBString(BytesStream, isc_dpb_user_name, aUserName);
    // Password
    Build_PBString(BytesStream, isc_dpb_password, aPassword);
    // 戻り値
    Result := BytesStream.Bytes;
    SetLength(Result, BytesStream.Position);
  finally
    BytesStream.Free;
  end;
end;

function Build_FileSpec(aHostName: string; aPort: uInt16; aDatabaseName: string): string;
// Filespec をビルド
begin
  Result := Trim(aHostName);
  if (Result <> '') and (aPort <> 0) and (aPort <> 3050) then
    Result := Result + Format('/%d', [aPort]);
  if (Result <> '') then
    Result := Result + ':';
  Result := Result + Trim(aDatabaseName);
end;

function Build_DatabaseName(aHostName: string; aPort: uInt16; aDatabaseName: string): TBytes;
// データベース文字列のビルド
var
  s: string;
begin
  s := Build_FileSpec(aHostName, aPort, aDatabaseName);
  Result := TEncoding.Convert(TEncoding.Default, TEncoding.ANSI, BytesOf(s));
end;

end.
