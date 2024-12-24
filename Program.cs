using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms; // Windowsフォームを使用
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

class ExtractTextByFontSize
{
    [STAThread]
    static void Main(string[] args)
    {
        // ファイル選択ダイアログを表示
        OpenFileDialog openFileDialog = new OpenFileDialog
        {
            Title = "PDFファイルを選択してください",
            Filter = "PDFファイル (*.pdf)|*.pdf",
            InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) // 初期フォルダをデスクトップに設定
        };

        if (openFileDialog.ShowDialog() == DialogResult.OK)
        {
            string pdfPath = openFileDialog.FileName;

            // 保存先フォルダ: ダウンロードフォルダ
            string downloadsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
            string outputFilePath = Path.Combine(downloadsPath, "GroupedByFontSizeText.txt");

            try
            {
                using (var pdf = PdfDocument.Open(pdfPath))
                {
                    using (StreamWriter writer = new StreamWriter(outputFilePath))
                    {
                        writer.WriteLine("--- PDFから抽出した内容 ---\n");

                        foreach (var page in pdf.GetPages())
                        {
                            writer.WriteLine($"ページ {page.Number}:\n");

                            // 文字をフォントサイズでグループ化
                            var groupedByFontSize = page.Letters
                                .GroupBy(l => Math.Round(l.FontSize, 1)) // フォントサイズでグループ化
                                .OrderByDescending(g => g.Key); // 大きいフォントサイズから順に処理

                            foreach (var fontSizeGroup in groupedByFontSize)
                            {
                                writer.WriteLine($"フォントサイズ: {fontSizeGroup.Key}\n");

                                // 各フォントサイズグループ内の文字をそのまま結合
                                string text = string.Join("", fontSizeGroup.Select(l => l.Value));
                                writer.WriteLine(text);

                                writer.WriteLine(); // フォントサイズグループの区切り
                            }

                            writer.WriteLine("\n-------------------\n");
                        }
                    }
                }

                MessageBox.Show($"PDFの内容を以下に保存しました:\n{outputFilePath}", "処理成功", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // メモ帳でファイルを開く
                System.Diagnostics.Process.Start("notepad.exe", outputFilePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"エラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        else
        {
            MessageBox.Show("ファイル選択がキャンセルされました。", "キャンセル", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }
}
