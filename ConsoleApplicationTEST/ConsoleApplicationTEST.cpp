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
#include <system/io/file.h>
#include <system/io/file_stream.h>
#include <system/io/path.h>
#include <iostream>
#include <chrono>
#include <thread>

using namespace Aspose::Words;

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

static void ApplyHeadingStyles(System::SharedPtr<Document> doc) {
    System::SharedPtr<NodeCollection> paragraphs = doc->GetChildNodes(NodeType::Paragraph, true);

    for (int i = 0; i < paragraphs->get_Count(); ++i) {
        System::SharedPtr<Paragraph> para = System::AsCast<Paragraph>(paragraphs->idx_get(i));
        if (!para) continue;

        System::SharedPtr<Style> paraStyle = System::AsCast<Style>(para->get_ParagraphFormat()->get_Style());
        if (!paraStyle) continue;

        // Проверяем, является ли абзац заголовком 1 или 2 уровня
        bool isHeading = (paraStyle->get_StyleIdentifier() == StyleIdentifier::Heading1 ||
            paraStyle->get_StyleIdentifier() == StyleIdentifier::Heading2);

        if (isHeading) {
            // Устанавливаем шрифт Times New Roman 14pt для заголовка
            para->get_ParagraphFormat()->get_Style()->get_Font()->set_Name(u"Times New Roman");
            para->get_ParagraphFormat()->get_Style()->get_Font()->set_Size(14);

            // Отступ сверху 24pt по умолчанию
            double spaceBefore = 24;
            double spaceAfter = 24;

            // Проверяем предыдущий абзац
            if (i > 0) {
                System::SharedPtr<Paragraph> prevPara = System::AsCast<Paragraph>(paragraphs->idx_get(i - 1));
                if (prevPara) {
                    System::SharedPtr<Style> prevParaStyle = System::AsCast<Style>(prevPara->get_ParagraphFormat()->get_Style());
                    if (prevParaStyle &&
                        (prevParaStyle->get_StyleIdentifier() == StyleIdentifier::Heading1 ||
                            prevParaStyle->get_StyleIdentifier() == StyleIdentifier::Heading2)) {
                        // Если предыдущий абзац - тоже заголовок, то отступ сверху 8pt
                        spaceBefore = 8;
                    }
                }
            }

            // Проверяем следующий абзац
            if (i + 1 < paragraphs->get_Count()) {
                System::SharedPtr<Paragraph> nextPara = System::AsCast<Paragraph>(paragraphs->idx_get(i + 1));
                if (nextPara) {
                    System::SharedPtr<Style> nextParaStyle = System::AsCast<Style>(nextPara->get_ParagraphFormat()->get_Style());
                    if (nextParaStyle &&
                        (nextParaStyle->get_StyleIdentifier() == StyleIdentifier::Heading1 ||
                            nextParaStyle->get_StyleIdentifier() == StyleIdentifier::Heading2)) {
                        // Если следующий абзац - тоже заголовок, то отступ снизу 8pt
                        spaceAfter = 8;
                    }
                }
            }

            // Применяем отступы
            para->get_ParagraphFormat()->set_SpaceBefore(spaceBefore);
            para->get_ParagraphFormat()->set_SpaceAfter(spaceAfter);

            // Принудительно устанавливаем шрифт для всех Run в заголовке
            System::SharedPtr<NodeCollection> runs = para->GetChildNodes(NodeType::Run, true);
            for (int j = 0; j < runs->get_Count(); ++j) {
                System::SharedPtr<Run> run = System::AsCast<Run>(runs->idx_get(j));
                if (run) {
                    run->get_Font()->set_Name(u"Times New Roman");
                    run->get_Font()->set_Size(14);
                }
            }
        }
    }
}

static void ApplyDocumentFormatting(System::SharedPtr<Document> doc) {
    // 1. Настройка стиля Times New Roman для всего документа
    System::SharedPtr<Style> style = System::AsCast<Style>(doc->get_Styles()->idx_get(u"Normal"));
    if (style) {
        style->get_Font()->set_Name(u"Times New Roman");
        style->get_Font()->set_Size(14);
    }

    // 2. Применяем стили к заголовкам
    ApplyHeadingStyles(doc);

    // 3. Настройка обычных абзацев
    System::SharedPtr<NodeCollection> paragraphs = doc->GetChildNodes(NodeType::Paragraph, true);
    for (int i = 0; i < paragraphs->get_Count(); ++i) {
        System::SharedPtr<Paragraph> para = System::AsCast<Paragraph>(paragraphs->idx_get(i));
        if (!para) continue;

        // Пропускаем заголовки (они уже обработаны в ApplyHeadingStyles)
        System::SharedPtr<Style> paraStyle = System::AsCast<Style>(para->get_ParagraphFormat()->get_Style());
        if (paraStyle && (paraStyle->get_StyleIdentifier() == StyleIdentifier::Heading1 ||
            paraStyle->get_StyleIdentifier() == StyleIdentifier::Heading2)) {
            continue;
        }

        // Настройка абзаца
        System::SharedPtr<ParagraphFormat> format = para->get_ParagraphFormat();
        format->set_FirstLineIndent(35.43);
        format->set_LineSpacingRule(LineSpacingRule::Multiple);
        format->set_LineSpacing(18);

        // Принудительная установка шрифта для всех Run
        System::SharedPtr<NodeCollection> runs = para->GetChildNodes(NodeType::Run, true);
        for (int j = 0; j < runs->get_Count(); ++j) {
            System::SharedPtr<Run> run = System::AsCast<Run>(runs->idx_get(j));
            if (run) {
                run->get_Font()->set_Name(u"Times New Roman");
                run->get_Font()->set_Size(14);
            }
        }
    }

    // 4. Настройка полей страницы
    System::SharedPtr<SectionCollection> sections = doc->get_Sections();
    for (int i = 1; i < sections->get_Count(); ++i) {
        System::SharedPtr<Section> section = System::AsCast<Section>(sections->idx_get(i));
        if (!section) continue;

        System::SharedPtr<PageSetup> pageSetup = section->get_PageSetup();
        pageSetup->set_LeftMargin(30);
        pageSetup->set_RightMargin(10);
        pageSetup->set_TopMargin(20);
        pageSetup->set_BottomMargin(20);
    }
}

int main() {
    System::String inputPath = u"D:\\Profiles\\Acer\\Desktop\\laba rest.docx";
    System::String tempPath = System::IO::Path::Combine(System::IO::Path::GetTempPath(), u"temp_laba.docx");
    System::String outputPath = u"D:\\Profiles\\Acer\\Desktop\\laba_processed.docx";

    try {
        // 1. Создаем временную копию файла
        System::IO::File::Copy(inputPath, tempPath, true);

        // 2. Работаем с копией
        System::SharedPtr<Document> doc;
        if (!TryOpenDocument(tempPath, doc)) {
            std::cerr << "Не удалось открыть файл после нескольких попыток." << std::endl;
            std::cerr << "Возможно, файл заблокирован другой программой." << std::endl;
            return 1;
        }

        // 3. Обработка документа
        ApplyDocumentFormatting(doc);

        // 4. Сохраняем результат
        doc->Save(outputPath);

        // 5. Удаляем временный файл
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