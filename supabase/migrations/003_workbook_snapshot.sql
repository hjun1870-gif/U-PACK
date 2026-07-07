-- 업로드 엑셀 파일을 DB에 보관 (모바일·다른 기기 복원용 — Storage 설정 없이도 동작)
CREATE TABLE IF NOT EXISTS workbook_snapshot (
  id TEXT PRIMARY KEY,
  file_base64 TEXT NOT NULL,
  updated_at TIMESTAMPTZ DEFAULT NOW()
);

ALTER TABLE workbook_snapshot ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Allow public read workbook_snapshot"
ON workbook_snapshot FOR SELECT USING (true);

CREATE POLICY "Allow public insert workbook_snapshot"
ON workbook_snapshot FOR INSERT WITH CHECK (true);

CREATE POLICY "Allow public update workbook_snapshot"
ON workbook_snapshot FOR UPDATE USING (true) WITH CHECK (true);

CREATE POLICY "Allow public delete workbook_snapshot"
ON workbook_snapshot FOR DELETE USING (true);
