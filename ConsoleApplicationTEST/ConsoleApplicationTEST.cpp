#include <Aspose.Words.Cpp/Document.h>
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
#include <system/io/file.h>
#include <system/io/file_stream.h>
#include <system/io/path.h>
#include <iostream>
#include <chrono>
#include <thread>

using namespace System;
using namespace System::IO;
using namespace Aspose::Words;
using namespace Aspose::Words::Tables;

static void ForceFormatHeadings(System::SharedPtr<Document> doc);
static void ForceFormatAllTables(System::SharedPtr<Document> doc);
static void FormatTitlePageText(System::SharedPtr<Document> doc);
static void ApplyDocumentFormatting(System::SharedPtr<Document> doc);
static bool TryOpenDocument(const System::String& filePath, System::SharedPtr<Document>& doc);

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

static void ForceFormatAllTables(System::SharedPtr<Document> doc) {
    System::SharedPtr<NodeCollection> tables = doc->GetChildNodes(NodeType::Table, true);

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
        para->get_ParagraphFormat()->set_FirstLineIndent(35.43); // Отступ первой строки (1.25 см)
        para->get_ParagraphFormat()->set_LineSpacingRule(LineSpacingRule::Multiple);
        para->get_ParagraphFormat()->set_LineSpacing(18); // Межстрочный интервал (1.5 строки)
    }

    // 4. Жесткое форматирование таблиц (оставляем как было)
    ForceFormatAllTables(doc);

    // 5. Настройка полей страниц (кроме титульной)
    for (int i = 1; i < doc->get_Sections()->get_Count(); ++i) {
        System::SharedPtr<Section> section = doc->get_Sections()->idx_get(i);
        System::SharedPtr<PageSetup> setup = section->get_PageSetup();
        setup->set_LeftMargin(30);   // 3 см (30 мм)
        setup->set_RightMargin(10);  // 1 см (10 мм)
        setup->set_TopMargin(20);    // 2 см (20 мм)
        setup->set_BottomMargin(20); // 2 см (20 мм)
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



int main() {
    setlocale(LC_ALL, "RU");
    System::String inputPath = u"D:\\Profiles\\Acer\\Desktop\\itogg.docx";
    System::String tempPath = System::IO::Path::Combine(System::IO::Path::GetTempPath(), u"temp_laba.docx");
    System::String outputPath = u"D:\\Profiles\\Acer\\Desktop\\laba_processed.docx";

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

        System::IO::File::Delete(tempPath);

        std::cout << "Документ успешно обработан и сохранен как: "
            << outputPath.ToUtf8String() << std::endl;
    }
    catch (const System::Exception& ex) {
        std::cerr << "Критическая ошибка: " << ex->get_Message().ToUtf8String() << std::endl;
        try {
            if (System::IO::File::Exists(tempPath)) {
                System::IO::File::Delete(tempPath);
            }
        }
        catch (...) {}
        return 1;
    }

    return 0;
}