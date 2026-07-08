-- 업로드한 엑셀 파일을 Supabase Storage에 보관 (새로고침·모바일 복원용)
INSERT INTO storage.buckets (id, name, public)
VALUES ('workbooks', 'workbooks', true)
ON CONFLICT (id) DO NOTHING;

CREATE POLICY "Allow public read workbooks"
ON storage.objects FOR SELECT
USING (bucket_id = 'workbooks');

CREATE POLICY "Allow public upload workbooks"
ON storage.objects FOR INSERT
WITH CHECK (bucket_id = 'workbooks');

CREATE POLICY "Allow public update workbooks"
ON storage.objects FOR UPDATE
USING (bucket_id = 'workbooks')
WITH CHECK (bucket_id = 'workbooks');

CREATE POLICY "Allow public delete workbooks"
ON storage.objects FOR DELETE
USING (bucket_id = 'workbooks');
