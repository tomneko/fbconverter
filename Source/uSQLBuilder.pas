{*******************************************************}
{                                                       }
{              Firebird Database Converter              }
{                                                       }
{*******************************************************}

// -----------------------------------------------------------------------------
// Note:
//   - SQL Build Utility Unit
// -----------------------------------------------------------------------------
unit uSQLBuilder;

interface

uses
  SysUtils;

type

  { TSelectSQL }
  TSelectSQL = record
    SELECT: string;
    FROM: string;
    WHERE: string;
    GROUP: string;
    HAVING: string;
    UNION: string;
    PLAN: string;
    ORDER: string;
    function Build(Terminate: Boolean = False): string;
    procedure Init;
  end;

  { TInsertSQL }
  TInsertSQL = record
    INSERT: string;
    FIELDS: string;
    VALUES: string;
    RETURNING: string;
    COMMENT: string;
    function Build(Terminate: Boolean = False): string;
    procedure Init;
  end;

  { TDeleteSQL }
  TDeleteSQL = record
    FROM: string;
    WHERE: string;
    function Build(Terminate: Boolean = False): string;
    procedure Init;
  end;

  { TUpdateSQL }
  TUpdateSQL = record
    UPDATE: string;
    &SET: string;
    WHERE: string;
    function Build(Terminate: Boolean = False): string;
    procedure Init;
  end;

const
  TERM_CHAR = ';';

implementation

const
  TerminaterStr: array [Boolean] of string = ('', TERM_CHAR);

{ TSelectSQL }

function TSelectSQL.Build(Terminate: Boolean): string;
// Select 文の生成
begin
  Result := Format('SELECT %s FROM %s', [Self.SELECT, Self.FROM]);
  if Self.WHERE <> '' then
    Result := Result + Format(' WHERE %s', [Self.WHERE]);
  if Self.GROUP <> '' then
    Result := Result + Format(' GROUP BY %s', [Self.GROUP]);
  if Self.HAVING <> '' then
    Result := Result + Format(' HAVING %s', [Self.HAVING]);
  if Self.UNION <> '' then
    Result := Result + Format(' UNION %s', [Self.UNION]);
  if Self.PLAN <> '' then
    Result := Result + Format(' PLAN %s', [Self.PLAN]);
  if Self.ORDER <> '' then
    Result := Result + Format(' ORDER BY %s', [Self.ORDER]);
  Result := Result + TerminaterStr[Terminate];
end;

procedure TSelectSQL.Init;
// Select 文のクリア
begin
  Self.SELECT := '';
  Self.FROM := '';
  Self.WHERE := '';
  Self.GROUP := '';
  Self.HAVING := '';
  Self.UNION := '';
  Self.PLAN := '';
  Self.ORDER := '';
end;

{ TInsertSQL }

function TInsertSQL.Build(Terminate: Boolean): string;
// Insert 文の生成
var
  dFields, dReturning: string;
begin
  if Self.FIELDS <> '' then
    dFields := Format(' (%s)', [Self.FIELDS])
  else
    dFields := '';
  if Self.RETURNING <> '' then
    dReturning := Format(' RETURNING %s', [Self.RETURNING])
  else
    dReturning := '';
  Result := Format('INSERT INTO %s%s VALUES (%s)%s%s%s',
                   [Self.INSERT, dFields, Self.VALUES, Self.RETURNING, TerminaterStr[Terminate], Self.COMMENT]);
end;

procedure TInsertSQL.Init;
// Insert 文のクリア
begin
  Self.INSERT := '';
  Self.FIELDS := '';
  Self.VALUES := '';
  Self.RETURNING := '';
  Self.COMMENT := '';
end;

{ TDeleteSQL }

function TDeleteSQL.Build(Terminate: Boolean): string;
// Delete 文の生成
begin
  Result := Format('DELETE FROM %s', [Self.FROM]);
  if Self.WHERE <> '' then
    Result := Result + Format(' WHERE %s', [Self.WHERE]);
  Result := Result + TerminaterStr[Terminate];
end;

procedure TDeleteSQL.Init;
// Delete 文のクリア
begin
  Self.FROM := '';
  Self.WHERE := '';
end;

{ TUpdateSQL }

function TUpdateSQL.Build(Terminate: Boolean): string;
begin
  Result := Format('UPDATE %s SET %s', [Self.UPDATE, Self.&SET]);
  if Self.WHERE <> '' then
    Result := Result + Format(' WHERE %s', [WHERE]);
  Result := Result + TerminaterStr[Terminate];
end;

procedure TUpdateSQL.Init;
// Update 文のクリア
begin
  Self.UPDATE := '';
  Self.&SET := '';
  Self.WHERE := '';
end;

end.
