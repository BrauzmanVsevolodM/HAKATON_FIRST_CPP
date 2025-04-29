#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/SaveFormat.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Font.h>
#include <Aspose.Words.Cpp/Style.h>
#include <Aspose.Words.Cpp/StyleIdentifier.h>
#include <Aspose.Words.Cpp/Run.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/SectionCollection.h>
#include <Aspose.Words.Cpp/StyleCollection.h>
#include <Aspose.Words.Cpp/PageSetup.h>
#include <Aspose.Words.Cpp/Tables/Table.h>
#include <Aspose.Words.Cpp/Tables/Row.h>
#include <Aspose.Words.Cpp/Tables/Cell.h>
#include <Aspose.Words.Cpp/Tables/CellFormat.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/RowCollection.h>
#include <Aspose.Words.Cpp/Model/Tables/CellCollection.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleCollection.h>
#include <Aspose.Words.Cpp/Layout/LayoutCollector.h>
#include <Aspose.Words.Cpp/Body.h>
#include <system/io/file.h>
#include <system/io/file_stream.h>
#include <system/io/path.h>
#include <iostream>
#include <chrono>
#include <thread>
#include <locale.h>
#include <system/string.h>
#include <fstream>




using namespace System;
using namespace System::IO;
using namespace Aspose::Words;
using namespace Aspose::Words::Tables;
using namespace Aspose::Words::Layout;

static void ForceFormatHeadings(System::SharedPtr<Document> doc);
static void ForceFormatAllTables(System::SharedPtr<Document> doc);
static void FormatTitlePageText(System::SharedPtr<Document> doc);
static void ApplyDocumentFormatting(System::SharedPtr<Document> doc);
static bool TryOpenDocument(const System::String& filePath, System::SharedPtr<Document>& doc);
static void CreateReportDocument(System::SharedPtr<Document> originalDoc, const System::String& reportPath);
static void WriteToReport(const System::String& reportPath, const System::String& message);
static void CheckDocumentFormatting(System::SharedPtr<Document> doc, const System::String& reportPath);
static void SaveAsPdf(const System::String& docxPath, const System::String& pdfPath);

static void ForceFormatHeadings(System::SharedPtr<Document> doc) {
    // 1. Обновляем стили заголовков
    {
        // Стиль Heading1
        System::SharedPtr<Style> heading1 = doc->get_Styles()->idx_get(StyleIdentifier::Heading1);
        if (heading1) {
            heading1->get_Font()->set_Name(u"Times New Roman");
            heading1->get_Font()->set_Size(14);
            heading1->get_Font()->set_Bold(false);
            heading1->get_ParagraphFormat()->set_SpaceBefore(24);
            heading1->get_ParagraphFormat()->set_SpaceAfter(24);
            heading1->get_ParagraphFormat()->set_LineSpacing(18);
        }

        // Стиль Heading2
        System::SharedPtr<Style> heading2 = doc->get_Styles()->idx_get(StyleIdentifier::Heading2);
        if (heading2) {
            heading2->get_Font()->set_Name(u"Times New Roman");
            heading2->get_Font()->set_Size(14);
            heading2->get_Font()->set_Bold(false);
            heading2->get_ParagraphFormat()->set_SpaceBefore(24);
            heading2->get_ParagraphFormat()->set_SpaceAfter(24);
            heading2->get_ParagraphFormat()->set_LineSpacing(18);
        }
    }

    // 2. Принудительно применяем стили ко всем заголовкам
    System::SharedPtr<NodeCollection> paragraphs = doc->GetChildNodes(NodeType::Paragraph, true);
    for (int i = 0; i < paragraphs->get_Count(); ++i) {
        System::SharedPtr<Paragraph> para = System::ExplicitCast<Paragraph>(paragraphs->idx_get(i));
        if (!para) continue;

        // Проверяем, является ли абзац заголовком
        System::SharedPtr<Style> paraStyle = para->get_ParagraphFormat()->get_Style();
        if (!paraStyle) continue;

        const bool isHeading1 = (paraStyle->get_StyleIdentifier() == StyleIdentifier::Heading1);
        const bool isHeading2 = (paraStyle->get_StyleIdentifier() == StyleIdentifier::Heading2);

        if (isHeading1 || isHeading2) {
            // Форматируем текст в заголовке
            System::SharedPtr<NodeCollection> runs = para->GetChildNodes(NodeType::Run, true);
            for (int j = 0; j < runs->get_Count(); ++j) {
                System::SharedPtr<Run> run = System::ExplicitCast<Run>(runs->idx_get(j));
                if (run) {
                    run->get_Font()->set_Name(u"Times New Roman");
                    run->get_Font()->set_Size(14);
                    run->get_Font()->set_Bold(false);
                }
            }

            // Настройки абзаца для заголовков
            para->get_ParagraphFormat()->set_SpaceBefore(24);
            para->get_ParagraphFormat()->set_SpaceAfter(24);
            para->get_ParagraphFormat()->set_LineSpacing(18);
        }
    }
}

static void SaveAsPdf(const System::String& docxPath, const System::String& pdfPath) {
    System::SharedPtr<Document> doc;
    if (!TryOpenDocument(docxPath, doc)) {
        throw System::Exception(u"Не удалось открыть документ для конвертации в PDF");
    }

    // Сохраняем в PDF
    doc->Save(pdfPath, SaveFormat::Pdf);
}

static void ForceFormatAllTables(System::SharedPtr<Document> doc) {
    System::SharedPtr<NodeCollection> tables = doc->GetChildNodes(NodeType::Table, true);

    for (int i = 0; i < tables->get_Count(); ++i) {
        System::SharedPtr<Table> table = System::ExplicitCast<Table>(tables->idx_get(i));
        if (!table) continue;

        // Выравниваем таблицу по центру страницы
        table->set_Alignment(TableAlignment::Center);

        for (int rowIdx = 0; rowIdx < table->get_Rows()->get_Count(); ++rowIdx) {
            System::SharedPtr<Row> row = table->get_Rows()->idx_get(rowIdx);
            if (!row) continue;

            for (int cellIdx = 0; cellIdx < row->get_Cells()->get_Count(); ++cellIdx) {
                System::SharedPtr<Cell> cell = row->get_Cells()->idx_get(cellIdx);
                if (!cell) continue;

                cell->get_CellFormat()->set_VerticalAlignment(CellVerticalAlignment::Center);

                System::SharedPtr<NodeCollection> paragraphs = cell->GetChildNodes(NodeType::Paragraph, true);
                for (int paraIdx = 0; paraIdx < paragraphs->get_Count(); ++paraIdx) {
                    System::SharedPtr<Paragraph> para = System::ExplicitCast<Paragraph>(paragraphs->idx_get(paraIdx));
                    if (!para) continue;

                    para->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Center);

                    System::SharedPtr<NodeCollection> runs = para->GetChildNodes(NodeType::Run, true);
                    for (int runIdx = 0; runIdx < runs->get_Count(); ++runIdx) {
                        System::SharedPtr<Run> run = System::ExplicitCast<Run>(runs->idx_get(runIdx));
                        if (run) {
                            run->get_Font()->set_Name(u"Times New Roman");
                            run->get_Font()->set_Size(10);
                            run->get_Font()->set_Bold(false);
                        }
                    }
                }
            }
        }
    }
}

static void FormatTitlePageText(System::SharedPtr<Document> doc) {
    System::SharedPtr<Section> titleSection = doc->get_FirstSection();
    if (!titleSection) return;

    // Обработка обычного текста
    System::SharedPtr<NodeCollection> runs = titleSection->GetChildNodes(NodeType::Run, true);
    for (int i = 0; i < runs->get_Count(); ++i) {
        System::SharedPtr<Run> run = System::ExplicitCast<Run>(runs->idx_get(i));
        if (run) {
            run->get_Font()->set_Name(u"Times New Roman");
            run->get_Font()->set_Size(14);
        }
    }

    // Обработка таблиц
    System::SharedPtr<NodeCollection> tables = titleSection->GetChildNodes(NodeType::Table, true);
    for (int i = 0; i < tables->get_Count(); ++i) {
        System::SharedPtr<Table> table = System::ExplicitCast<Table>(tables->idx_get(i));
        if (!table) continue;

        for (int rowIdx = 0; rowIdx < table->get_Rows()->get_Count(); ++rowIdx) {
            System::SharedPtr<Row> row = table->get_Rows()->idx_get(rowIdx);
            if (!row) continue;

            for (int cellIdx = 0; cellIdx < row->get_Cells()->get_Count(); ++cellIdx) {
                System::SharedPtr<Cell> cell = row->get_Cells()->idx_get(cellIdx);
                if (!cell) continue;

                System::SharedPtr<NodeCollection> paragraphs = cell->GetChildNodes(NodeType::Paragraph, true);
                for (int paraIdx = 0; paraIdx < paragraphs->get_Count(); ++paraIdx) {
                    System::SharedPtr<Paragraph> para = System::ExplicitCast<Paragraph>(paragraphs->idx_get(paraIdx));
                    if (!para) continue;

                    para->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Left);

                    System::SharedPtr<NodeCollection> paraRuns = para->GetChildNodes(NodeType::Run, true);
                    for (int runIdx = 0; runIdx < paraRuns->get_Count(); ++runIdx) {
                        System::SharedPtr<Run> run = System::ExplicitCast<Run>(paraRuns->idx_get(runIdx));
                        if (run) {
                            run->get_Font()->set_Name(u"Times New Roman");
                            run->get_Font()->set_Size(14);
                            run->get_Font()->set_Bold(false);
                        }
                    }
                }
            }
        }
    }
}

static void ApplyDocumentFormatting(System::SharedPtr<Document> doc) {
    // 1. Жесткое форматирование заголовков
    ForceFormatHeadings(doc);

    // 2. Установка базового стиля Normal (Times New Roman, 14pt)
    System::SharedPtr<Style> normalStyle = doc->get_Styles()->idx_get(u"Normal");
    if (normalStyle) {
        normalStyle->get_Font()->set_Name(u"Times New Roman");
        normalStyle->get_Font()->set_Size(14);
    }

    // 3. Форматирование обычного текста
    System::SharedPtr<NodeCollection> paragraphs = doc->GetChildNodes(NodeType::Paragraph, true);
    for (int i = 0; i < paragraphs->get_Count(); ++i) {
        System::SharedPtr<Paragraph> para = System::ExplicitCast<Paragraph>(paragraphs->idx_get(i));
        if (!para) continue;

        // Пропускаем заголовки и таблицы
        System::SharedPtr<Style> paraStyle = para->get_ParagraphFormat()->get_Style();
        if (paraStyle && (paraStyle->get_StyleIdentifier() == StyleIdentifier::Heading1 ||
            paraStyle->get_StyleIdentifier() == StyleIdentifier::Heading2)) {
            continue;
        }
        if (para->GetAncestor(NodeType::Table) != nullptr) continue;

        // Принудительно применяем Times New Roman 14pt ко всему тексту
        System::SharedPtr<NodeCollection> runs = para->GetChildNodes(NodeType::Run, true);
        for (int j = 0; j < runs->get_Count(); ++j) {
            System::SharedPtr<Run> run = System::ExplicitCast<Run>(runs->idx_get(j));
            if (run) {
                run->get_Font()->set_Name(u"Times New Roman");
                run->get_Font()->set_Size(14);
                run->get_Font()->set_Bold(false);
                run->get_Font()->set_Italic(false);
            }
        }

        // Форматирование абзаца (отступы и межстрочный интервал)
        para->get_ParagraphFormat()->set_FirstLineIndent(35.43);
        para->get_ParagraphFormat()->set_LineSpacingRule(LineSpacingRule::Multiple);
        para->get_ParagraphFormat()->set_LineSpacing(18);
    }

    // 4. Жесткое форматирование таблиц (оставляем как было)
    ForceFormatAllTables(doc);

    // 5. Настройка полей страниц (кроме титульной)
    for (int i = 1; i < doc->get_Sections()->get_Count(); ++i) {
        System::SharedPtr<Section> section = doc->get_Sections()->idx_get(i);
        System::SharedPtr<PageSetup> setup = section->get_PageSetup();
        setup->set_LeftMargin(30);
        setup->set_RightMargin(10);
        setup->set_TopMargin(20);
        setup->set_BottomMargin(20);
    }
}

static bool TryOpenDocument(const System::String& filePath, System::SharedPtr<Document>& doc) {
    const int maxAttempts = 5;
    const int delayMs = 500;

    for (int attempt = 1; attempt <= maxAttempts; ++attempt) {
        try {
            System::SharedPtr<System::IO::FileStream> stream = System::MakeObject<System::IO::FileStream>(
                filePath, System::IO::FileMode::Open, System::IO::FileAccess::Read);
            doc = System::MakeObject<Document>(stream);
            return true;
        }
        catch (const System::IO::IOException&) {
            if (attempt == maxAttempts) return false;
            std::this_thread::sleep_for(std::chrono::milliseconds(delayMs));
        }
    }
    return false;
}

static void WriteToReport(const System::String& reportPath, const System::String& message) {
    try {
        std::ofstream outFile;
        outFile.open(reportPath.ToUtf8String(), std::ios::app);
        if (outFile.is_open()) {
            outFile << message.ToUtf8String() << std::endl;
            outFile.close();
        }
    }
    catch (...) {
        std::cerr << "Error writing to report file" << std::endl;
    }
}

static void CheckDocumentFormatting(System::SharedPtr<Document> doc, const System::String& reportPath) {
    System::SharedPtr<LayoutCollector> layoutCollector = System::MakeObject<LayoutCollector>(doc);

    if (File::Exists(reportPath)) {
        File::Delete(reportPath);
    }

    WriteToReport(reportPath, u"ОТЧЕТ ОБ ОШИБКАХ ФОРМАТИРОВАНИЯ");
    WriteToReport(reportPath, u"=================================");

    // ==================== ПРОВЕРКА ЗАГОЛОВКОВ ====================
    System::SharedPtr<NodeCollection> paragraphs = doc->GetChildNodes(NodeType::Paragraph, true);
    for (int i = 0; i < paragraphs->get_Count(); ++i) {
System::SharedPtr<Paragraph> para = System::AsCast<Paragraph>(paragraphs->idx_get(i));
        if (!para) continue;

        int pageNum = layoutCollector->GetStartPageIndex(para) + 1;
        System::String paraText = para->GetText().Trim();

        // Проверка стилей заголовков
        if (para->get_ParagraphFormat()->get_Style() != nullptr) {
            System::SharedPtr<Style> style = para->get_ParagraphFormat()->get_Style();

            if (style->get_StyleIdentifier() == StyleIdentifier::Heading1 ||
                style->get_StyleIdentifier() == StyleIdentifier::Heading2) {

                // Проверка шрифта заголовков
                if (style->get_Font()->get_Name() != u"Times New Roman") {
                    System::String error = System::String::Format(
                        u"Страница {0}: Заголовок '{1}' - неправильный шрифт (должен быть Times New Roman)",
                        pageNum, paraText);
                    WriteToReport(reportPath, error);
                }

                // Проверка размера шрифта заголовков
                if (style->get_Font()->get_Size() != 14) {
                    System::String error = System::String::Format(
                        u"Страница {0}: Заголовок '{1}' - неправильный размер шрифта (должен быть 14pt)",
                        pageNum, paraText);
                    WriteToReport(reportPath, error);
                }

                // Проверка жирности заголовков
                if (style->get_Font()->get_Bold()) {
                    System::String error = System::String::Format(
                        u"Страница {0}: Заголовок '{1}' - не должен быть жирным",
                        pageNum, paraText);
                    WriteToReport(reportPath, error);
                }

                // Проверка отступов заголовков
                if (style->get_ParagraphFormat()->get_SpaceBefore() != 24 ||
                    style->get_ParagraphFormat()->get_SpaceAfter() != 24) {
                    System::String error = System::String::Format(
                        u"Страница {0}: Заголовок '{1}' - неправильные отступы (должны быть 24pt до и после)",
                        pageNum, paraText);
                    WriteToReport(reportPath, error);
                }

                // Проверка межстрочного интервала заголовков
                if (style->get_ParagraphFormat()->get_LineSpacing() != 18) {
                    System::String error = System::String::Format(
                        u"Страница {0}: Заголовок '{1}' - неправильный межстрочный интервал (должен быть 18pt)",
                        pageNum, paraText);
                    WriteToReport(reportPath, error);
                }
            }
        }

        // ==================== ПРОВЕРКА ОСНОВНОГО ТЕКСТА ====================
        if (para->GetAncestor(NodeType::Table) == nullptr) {
            // Проверка отступа первой строки
            if (para->get_ParagraphFormat()->get_FirstLineIndent() != 35.43) {
                System::String error = System::String::Format(
                    u"Страница {0}: Абзац - неправильный отступ первой строки (должен быть 1.25 см)",
                    pageNum);
                WriteToReport(reportPath, error);
            }

            // Проверка межстрочного интервала
            if (para->get_ParagraphFormat()->get_LineSpacing() != 18) {
                System::String error = System::String::Format(
                    u"Страница {0}: Абзац - неправильный межстрочный интервал (должен быть 1.0)",
                    pageNum);
                WriteToReport(reportPath, error);
            }
        }

        // Проверка текста в абзацах
        System::SharedPtr<NodeCollection> runs = para->GetChildNodes(NodeType::Run, true);
        for (int j = 0; j < runs->get_Count(); ++j) {
            System::SharedPtr<Run> run = System::AsCast<Run>(runs->idx_get(j));
            if (!run) continue;

            // Проверка шрифта основного текста
            if (run->get_Font()->get_Name() != u"Times New Roman") {
                System::String error = System::String::Format(
                    u"Страница {0}: Обнаружен текст с неправильным шрифтом '{1}' (должен быть Times New Roman)",
                    pageNum, run->get_Font()->get_Name());
                WriteToReport(reportPath, error);
            }

            // Проверка размера шрифта
            if (run->get_Font()->get_Size() != 14) {
                System::String error = System::String::Format(
                    u"Страница {0}: Обнаружен текст с неправильным размером шрифта '{1}' (должен быть 14pt)",
                    pageNum, run->get_Font()->get_Size());
                WriteToReport(reportPath, error);
            }

            // Проверка жирности текста
            if (run->get_Font()->get_Bold()) {
                System::String error = System::String::Format(
                    u"Страница {0}: Обнаружен жирный текст (не должен быть жирным)",
                    pageNum);
                WriteToReport(reportPath, error);
            }

            // Проверка курсива
            if (run->get_Font()->get_Italic()) {
                System::String error = System::String::Format(
                    u"Страница {0}: Обнаружен курсивный текст (не должен быть курсивом)",
                    pageNum);
                WriteToReport(reportPath, error);
            }
        }
    }

    // ==================== ПРОВЕРКА ТАБЛИЦ ====================
    System::SharedPtr<NodeCollection> tables = doc->GetChildNodes(NodeType::Table, true);
    for (int i = 0; i < tables->get_Count(); ++i) {
        System::SharedPtr<Table> table = System::AsCast<Table>(tables->idx_get(i));
        if (!table) continue;

        int pageNum = layoutCollector->GetStartPageIndex(table) + 1;

        // Проверка выравнивания таблицы
        if (table->get_Alignment() != TableAlignment::Center) {
            System::String error = System::String::Format(
                u"Страница {0}: Таблица {1} - не выровнена по центру",
                pageNum, i + 1);
            WriteToReport(reportPath, error);
        }

        // Проверка содержимого таблицы
        for (int rowIdx = 0; rowIdx < table->get_Rows()->get_Count(); ++rowIdx) {
            System::SharedPtr<Row> row = table->get_Rows()->idx_get(rowIdx);
            if (!row) continue;

            for (int cellIdx = 0; cellIdx < row->get_Cells()->get_Count(); ++cellIdx) {
                System::SharedPtr<Cell> cell = row->get_Cells()->idx_get(cellIdx);
                if (!cell) continue;

                // Проверка вертикального выравнивания
                if (cell->get_CellFormat()->get_VerticalAlignment() != CellVerticalAlignment::Center) {
                    System::String error = System::String::Format(
                        u"Страница {0}: Ячейка ({1},{2}) - не выровнена по вертикали",
                        pageNum, rowIdx + 1, cellIdx + 1);
                    WriteToReport(reportPath, error);
                }

                // Проверка текста в ячейках
                System::SharedPtr<NodeCollection> cellParagraphs = cell->GetChildNodes(NodeType::Paragraph, true);
                for (int paraIdx = 0; paraIdx < cellParagraphs->get_Count(); ++paraIdx) {
                    System::SharedPtr<Paragraph> para = System::AsCast<Paragraph>(cellParagraphs->idx_get(paraIdx));
                    if (!para) continue;

                    // Проверка выравнивания текста
                    if (para->get_ParagraphFormat()->get_Alignment() != ParagraphAlignment::Center) {
                        System::String error = System::String::Format(
                            u"Страница {0}: Текст в ячейке ({1},{2}) - не выровнен по центру",
                            pageNum, rowIdx + 1, cellIdx + 1);
                        WriteToReport(reportPath, error);
                    }

                    // Проверка шрифта в таблицах
                    System::SharedPtr<NodeCollection> runs = para->GetChildNodes(NodeType::Run, true);
                    for (int runIdx = 0; runIdx < runs->get_Count(); ++runIdx) {
                        System::SharedPtr<Run> run = System::AsCast<Run>(runs->idx_get(runIdx));
                        if (!run) continue;

                        if (run->get_Font()->get_Name() != u"Times New Roman") {
                            System::String error = System::String::Format(
                                u"Страница {0}: Текст в ячейке ({1},{2}) - неправильный шрифт",
                                pageNum, rowIdx + 1, cellIdx + 1);
                            WriteToReport(reportPath, error);
                        }

                        if (run->get_Font()->get_Size() != 10) {
                            System::String error = System::String::Format(
                                u"Страница {0}: Текст в ячейке ({1},{2}) - неправильный размер шрифта (должен быть 10pt)",
                                pageNum, rowIdx + 1, cellIdx + 1);
                            WriteToReport(reportPath, error);
                        }
                    }
                }
            }
        }
    }

    // ==================== ПРОВЕРКА ПОЛЕЙ СТРАНИЦ ====================
    for (int i = 0; i < doc->get_Sections()->get_Count(); ++i) {
        System::SharedPtr<Section> section = doc->get_Sections()->idx_get(i);
        System::SharedPtr<PageSetup> pageSetup = section->get_PageSetup();
        int pageNum = layoutCollector->GetStartPageIndex(section->get_Body()->get_FirstParagraph()) + 1;

        if (i > 0) { // Для всех страниц кроме титульной
            if (pageSetup->get_LeftMargin() != 30) {
                System::String error = System::String::Format(
                    u"Страница {0}: Левое поле должно быть 30pt, получено {1}",
                    pageNum, pageSetup->get_LeftMargin());
                WriteToReport(reportPath, error);
            }

            if (pageSetup->get_RightMargin() != 10) {
                System::String error = System::String::Format(
                    u"Страница {0}: Правое поле должно быть 10pt, получено {1}",
                    pageNum, pageSetup->get_RightMargin());
                WriteToReport(reportPath, error);
            }

            if (pageSetup->get_TopMargin() != 20) {
                System::String error = System::String::Format(
                    u"Страница {0}: Верхнее поле должно быть 20pt, получено {1}",
                    pageNum, pageSetup->get_TopMargin());
                WriteToReport(reportPath, error);
            }

            if (pageSetup->get_BottomMargin() != 20) {
                System::String error = System::String::Format(
                    u"Страница {0}: Нижнее поле должно быть 20pt, получено {1}",
                    pageNum, pageSetup->get_BottomMargin());
                WriteToReport(reportPath, error);
            }
        }
    }

    WriteToReport(reportPath, u"=================================");
    WriteToReport(reportPath, u"Проверка завершена");
}

static void CreateReportDocument(System::SharedPtr<Document> originalDoc, const System::String& reportPath) {
    // Создаем новый документ для отчета
    System::SharedPtr<Document> reportDoc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(reportDoc);

    // Добавляем заголовок отчета
    builder->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Center);
    builder->get_Font()->set_Name(u"Times New Roman");
    builder->get_Font()->set_Size(14);
    builder->get_Font()->set_Bold(true);
    builder->Writeln(u"ОТЧЕТ О ПРОВЕРКЕ ФОРМАТИРОВАНИЯ ДОКУМЕНТА");
    builder->Writeln();

    // Проверяем форматирование исходного документа
    System::String tempReportPath = System::IO::Path::Combine(System::IO::Path::GetTempPath(), u"temp_report.txt");
    CheckDocumentFormatting(originalDoc, tempReportPath);

    // Читаем временный отчет и добавляем его в документ
    builder->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Left);
    builder->get_Font()->set_Bold(false);

    try {
        System::SharedPtr<System::IO::StreamReader> reader = System::MakeObject<System::IO::StreamReader>(tempReportPath);
        System::String line;
        while ((line = reader->ReadLine()) != nullptr) {
            builder->Writeln(line);
        }
        reader->Close();
    }
    catch (...) {
        builder->Writeln(u"Не удалось прочитать отчет о проверке форматирования");
    }

    // Сохраняем отчет
    reportDoc->Save(reportPath);
    System::IO::File::Delete(tempReportPath);
}

int main() {
    setlocale(LC_ALL, "Russian");
    System::String inputPath = u"D:\\Profiles\\Acer\\Desktop\\itogg.docx";
    System::String tempPath = System::IO::Path::Combine(System::IO::Path::GetTempPath(), u"temp_doc.docx");
    System::String outputPath = u"D:\\Profiles\\Acer\\Desktop\\laba_processed.docx";
    System::String pdfOutputPath = u"D:\\Profiles\\Acer\\Desktop\\laba_processed.pdf";
    System::String reportPath = u"D:\\Profiles\\Acer\\Desktop\\report.txt";

    try {
        System::IO::File::Copy(inputPath, tempPath, true);

        System::SharedPtr<Document> doc;
        if (!TryOpenDocument(tempPath, doc)) {
            std::cerr << "Не удалось открыть файл после нескольких попыток." << std::endl;
            std::cerr << "Возможно, файл заблокирован другой программой." << std::endl;
            return 1;
        }

        ApplyDocumentFormatting(doc);
        doc->Save(outputPath);
        CheckDocumentFormatting(doc, reportPath);
        doc->Save(pdfOutputPath, SaveFormat::Pdf);
        System::IO::File::Delete(tempPath);

        // Правильный вывод строк без ->
        std::cout << "DOCX: " << outputPath.ToUtf8String() << std::endl
            << "PDF: " << pdfOutputPath.ToUtf8String() << std::endl
            << "Отчет: " << reportPath.ToUtf8String() << std::endl;
    }
    catch (const System::Exception& ex) {
        try {
            if (System::IO::File::Exists(tempPath)) {
                System::IO::File::Delete(tempPath);
            }
        }
        catch (...) {}

        // Универсальный способ вывода сообщения об ошибке
        std::cerr << "Ошибка: ";
        try {
            // Попробуем разные варианты доступа к сообщению
            std::cerr << "Ошибка: " << ex->get_Message().ToUtf8String() << std::endl;
        }
        catch (...) {
            try {
                    
            }
            catch (...) {
                std::cerr << "Неизвестная ошибка";
            }
        }
        std::cerr << std::endl;
        return 1;
    }

    return 0;
}