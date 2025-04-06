#include <iostream>
#include <windows.h>
#include <comdef.h>
#import "C:\Program Files (x86)\Microsoft Office\root\Office16\MSWORD.OLB" rename("ExitWindows", "WordExitWindows")
#import "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
#import "C:\Program Files (x86)\Microsoft Office\root\Office16\MSPPT.OLB"

using namespace std;

// DeepSeek API 交互函数
string callDeepSeekAPI(const string& content) {
    // 这里实现与DeepSeek API的交互
    // 示例代码 - 需要替换为实际的API调用
    return "DeepSeek response for: " + content;
}

// 处理Word文档
void processWordDocument(const string& filePath) {
    CoInitialize(NULL);
    try {
        Word::_ApplicationPtr wordApp;
        wordApp.CreateInstance(__uuidof(Word::Application));
        wordApp->Visible = false;

        Word::_DocumentPtr doc = wordApp->Documents->Open(_bstr_t(filePath.c_str()));
        
        // 获取文档内容
        string content = (char*)doc->Content->Text;
        
        // 调用DeepSeek API
        string response = callDeepSeekAPI(content);
        cout << "Word Document Processed. Response: " << response << endl;
        
        doc->Close(false);
        wordApp->Quit();
    } catch (_com_error &e) {
        cerr << "Word Error: " << e.ErrorMessage() << endl;
    }
    CoUninitialize();
}

// 处理Excel文档
void processExcelDocument(const string& filePath) {
    CoInitialize(NULL);
    try {
        Excel::_ApplicationPtr excelApp;
        excelApp.CreateInstance(__uuidof(Excel::Application));
        excelApp->Visible = false;

        Excel::_WorkbookPtr workbook = excelApp->Workbooks->Open(_bstr_t(filePath.c_str()));
        
        // 获取第一个工作表的内容
        Excel::_WorksheetPtr sheet = workbook->Worksheets->GetItem(1);
        Excel::RangePtr usedRange = sheet->UsedRange;
        
        // 获取所有单元格内容
        _variant_t values = usedRange->Value2;
        
        // 转换为字符串
        string content;
        SAFEARRAY* sa = values.parray;
        long rows = sa->rgsabound[0].cElements;
        long cols = sa->rgsabound[1].cElements;
        
        for (long i = 1; i <= rows; ++i) {
            for (long j = 1; j <= cols; ++j) {
                _variant_t cellValue;
                cellValue.vt = VT_EMPTY;
                long indices[2] = {i-1, j-1};
                SafeArrayGetElement(sa, indices, &cellValue);
                
                if (cellValue.vt != VT_EMPTY && cellValue.vt != VT_NULL) {
                    _bstr_t bstrVal(cellValue);
                    content += (char*)bstrVal + string("\t");
                }
            }
            content += "\n";
        }
        
        // 调用DeepSeek API
        string response = callDeepSeekAPI(content);
        cout << "Excel Document Processed. Response: " << response << endl;
        
        workbook->Close(false);
        excelApp->Quit();
    } catch (_com_error &e) {
        cerr << "Excel Error: " << e.ErrorMessage() << endl;
    }
    CoUninitialize();
}

// 处理PowerPoint文档
void processPowerPointDocument(const string& filePath) {
    CoInitialize(NULL);
    try {
        PowerPoint::_ApplicationPtr pptApp;
        pptApp.CreateInstance(__uuidof(PowerPoint::Application));
        pptApp->Visible = false;

        PowerPoint::_PresentationPtr pres = pptApp->Presentations->Open(_bstr_t(filePath.c_str()));
        
        string content;
        for (int i = 1; i <= pres->Slides->Count; ++i) {
            PowerPoint::_SlidePtr slide = pres->Slides->Item(i);
            
            for (int j = 1; j <= slide->Shapes->Count; ++j) {
                PowerPoint::ShapePtr shape = slide->Shapes->Item(j);
                if (shape->HasTextFrame == PowerPoint::MsoTriState::msoTrue) {
                    if (shape->TextFrame->HasText == PowerPoint::MsoTriState::msoTrue) {
                        _bstr_t text = shape->TextFrame->TextRange->Text;
                        content += (char*)text + "\n";
                    }
                }
            }
        }
        
        // 调用DeepSeek API
        string response = callDeepSeekAPI(content);
        cout << "PowerPoint Document Processed. Response: " << response << endl;
        
        pres->Close();
        pptApp->Quit();
    } catch (_com_error &e) {
        cerr << "PowerPoint Error: " << e.ErrorMessage() << endl;
    }
    CoUninitialize();
}

int main() {
    cout << "Office to DeepSeek Connector" << endl;
    cout << "1. Process Word Document" << endl;
    cout << "2. Process Excel Document" << endl;
    cout << "3. Process PowerPoint Document" << endl;
    cout << "Enter choice (1-3): ";
    
    int choice;
    cin >> choice;
    
    cout << "Enter file path: ";
    string filePath;
    cin >> filePath;
    
    switch(choice) {
        case 1:
            processWordDocument(filePath);
            break;
        case 2:
            processExcelDocument(filePath);
            break;
        case 3:
            processPowerPointDocument(filePath);
            break;
        default:
            cerr << "Invalid choice" << endl;
    }
    
    return 0;
}