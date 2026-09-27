import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Styles definition
FONT_TITLE = Font(name="Calibri", size=14, bold=True, color="1F4E79")
FONT_HEADER = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
FONT_DATA = Font(name="Calibri", size=11, color="000000")
FONT_DATA_BOLD = Font(name="Calibri", size=11, bold=True, color="000000")
FONT_STRUCK = Font(name="Calibri", size=11, strike=True, color="7F7F7F")

FILL_HEADER = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
FILL_ZEBRA = PatternFill(start_color="F7F9FB", end_color="F7F9FB", fill_type="solid")
FILL_WHITE = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
FILL_STRUCK = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

BORDER_THIN = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)

ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center", wrap_text=True)

HEADERS = [
    "क्र.सं. (S.No.)",
    "प्रखंड (Block)",
    "पंचायत / स्थान (Panchayat / Location)",
    "पद / प्रभार (Designation / Role)",
    "नाम (Name)",
    "पिता / अभिभावक का नाम (Father's Name)",
    "मोबाइल नंबर (Mobile No.)",
    "उम्र (Age)",
    "श्रेणी (Category)",
    "जाति (Caste)",
    "संक्षिप्त राजनीतिक एवं सामाजिक विवरण (Profile / Details)",
    "कॉलिंग स्थिति (Calling Status)",
    "टिप्पणी (Remarks)",
    "स्रोत फ़ाइल (Source File)"
]

# 6 Newly added blocks from 5 pages + 1 PDF (Purnia District)
purnia_blocks = [
    {
        "sheet_title": "धमदाहा (Dhamdaha)",
        "block_name": "धमदाहा (Dhamdaha)",
        "district": "पूर्णियाँ (Purnia)",
        "source": " .pdf",
        "records": [
            {
                "sno": 1,
                "block": "धमदाहा",
                "panchayat": "",
                "role": "प्रखंड प्रभारी",
                "name": "ई० राहुल कु० (समर)",
                "father": "श्री नारायण मंडल",
                "mobile": "9570611977",
                "age": 33,
                "category": "OBC / EBC",
                "caste": "धानुक",
                "profile": "",
                "calling_status": "",
                "remarks": "प्रखंड प्रभारी",
                "struck": False
            },
            {
                "sno": 2,
                "block": "धमदाहा",
                "panchayat": "",
                "role": "प्रखंड प्रभारी",
                "name": "खगेश कु० सिंह",
                "father": "स्व० परमानन्द सिंह",
                "mobile": "7033350963",
                "age": 48,
                "category": "General",
                "caste": "राजपूत",
                "profile": "",
                "calling_status": "",
                "remarks": "प्रखंड प्रभारी",
                "struck": False
            },
            {
                "sno": 3,
                "block": "धमदाहा",
                "panchayat": "",
                "role": "प्रखंड चुनाव अभियान समिति सदस्य",
                "name": "विजय कु० साह",
                "father": "जोगेश्वर साह",
                "mobile": "7488541496",
                "age": 30,
                "category": "OBC",
                "caste": "तेली",
                "profile": "",
                "calling_status": "",
                "remarks": "नाम व पिता के नाम में 'सिंह' काटकर 'साह' लिखा गया",
                "struck": False
            },
            {
                "sno": 4,
                "block": "धमदाहा",
                "panchayat": "",
                "role": "सदस्य",
                "name": "अरविन्द कुमार",
                "father": "कामेश्वर मेहता",
                "mobile": "8877894018",
                "age": 40,
                "category": "OBC",
                "caste": "कुशवाहा",
                "profile": "",
                "calling_status": "",
                "remarks": "",
                "struck": False
            },
            {
                "sno": 5,
                "block": "धमदाहा",
                "panchayat": "",
                "role": "सदस्य",
                "name": "अभिनेन्दन कुमार",
                "father": "कन्हैया मिश्रा",
                "mobile": "8789179702",
                "age": 29,
                "category": "OBC / EBC",
                "caste": "हलवाई",
                "profile": "",
                "calling_status": "",
                "remarks": "",
                "struck": False
            },
            {
                "sno": 6,
                "block": "धमदाहा",
                "panchayat": "",
                "role": "सदस्य",
                "name": "रिजवान आलम",
                "father": "हासिम अली",
                "mobile": "9934563790",
                "age": 39,
                "category": "Minority",
                "caste": "मुस्लिम",
                "profile": "",
                "calling_status": "",
                "remarks": "",
                "struck": False
            },
            {
                "sno": 7,
                "block": "धमदाहा",
                "panchayat": "",
                "role": "सदस्य",
                "name": "विष्णुदेव रजक",
                "father": "रेवामलाल रजक",
                "mobile": "8002171060",
                "age": 62,
                "category": "SC",
                "caste": "धोबी",
                "profile": "",
                "calling_status": "",
                "remarks": "",
                "struck": False
            },
            {
                "sno": 8,
                "block": "धमदाहा",
                "panchayat": "",
                "role": "सदस्य",
                "name": "कुन्दन कु० दास",
                "father": "बैजनाथ दास",
                "mobile": "8102642701",
                "age": 35,
                "category": "OBC",
                "caste": "बनिया",
                "profile": "",
                "calling_status": "",
                "remarks": "मो० 8102642701",
                "struck": False
            },
            {
                "sno": 9,
                "block": "धमदाहा",
                "panchayat": "",
                "role": "सदस्य",
                "name": "गुड्डू कुमार",
                "father": "जय कुमार ठाकुर",
                "mobile": "8825265852",
                "age": 29,
                "category": "EBC",
                "caste": "नाई",
                "profile": "",
                "calling_status": "",
                "remarks": "",
                "struck": False
            },
            {
                "sno": 10,
                "block": "धमदाहा",
                "panchayat": "",
                "role": "सदस्य",
                "name": "चंदन कुमार",
                "father": "बिलो मंडल",
                "mobile": "7488008045",
                "age": 30,
                "category": "EBC",
                "caste": "कैवर्त",
                "profile": "",
                "calling_status": "",
                "remarks": "",
                "struck": False
            },
            {
                "sno": 11,
                "block": "धमदाहा",
                "panchayat": "",
                "role": "सदस्य",
                "name": "भरत पासवान",
                "father": "फौजी पासवान",
                "mobile": "9934815855",
                "age": 50,
                "category": "SC",
                "caste": "पासवान",
                "profile": "",
                "calling_status": "",
                "remarks": "",
                "struck": False
            },
            {
                "sno": 12,
                "block": "धमदाहा",
                "panchayat": "",
                "role": "सदस्य",
                "name": "अमरजीत कुमार",
                "father": "नंदकिशोर यादव",
                "mobile": "8406858531",
                "age": 29,
                "category": "OBC",
                "caste": "यादव",
                "profile": "",
                "calling_status": "",
                "remarks": "",
                "struck": False
            },
            {
                "sno": 13,
                "block": "धमदाहा",
                "panchayat": "",
                "role": "सदस्य",
                "name": "निर्भय कु० सिंह",
                "father": "अनिरुद्ध कु० सिंह",
                "mobile": "9955231452",
                "age": 55,
                "category": "General",
                "caste": "राजपूत",
                "profile": "",
                "calling_status": "",
                "remarks": "",
                "struck": False
            }
        ]
    },
    {
        "sheet_title": "भवानीपुर (Bhawanipur)",
        "block_name": "भवानीपुर (Bhawanipur)",
        "district": "पूर्णियाँ (Purnia)",
        "source": "WhatsApp Image 2026-09-25 at 4.52.41 PM.jpeg",
        "records": [
            {
                "sno": 1,
                "block": "भवानीपुर",
                "panchayat": "",
                "role": "प्रखंड प्रभारी",
                "name": "सदन कु० भगत",
                "father": "स्व० रामानंद भगत",
                "mobile": "6201154997",
                "age": 46,
                "category": "OBC",
                "caste": "कलवार",
                "profile": "जिला परिषद पुत्र राजनीति",
                "calling_status": "",
                "remarks": "प्रखंड प्रभारी | जिला परिषद पुत्र",
                "struck": False
            },
            {
                "sno": 2,
                "block": "भवानीपुर",
                "panchayat": "",
                "role": "प्रखंड प्रभारी",
                "name": "दिलशाद आलम",
                "father": "मो० मोहसिन",
                "mobile": "7370866254",
                "age": 32,
                "category": "Minority",
                "caste": "शेख",
                "profile": "सरपंच पुत्र पंचायती राजनीति",
                "calling_status": "",
                "remarks": "प्रखंड प्रभारी | सरपंच पुत्र",
                "struck": False
            },
            {
                "sno": 3,
                "block": "भवानीपुर",
                "panchayat": "",
                "role": "प्रखंड चुनाव अभियान समिति सदस्य",
                "name": "मिथुन कु० साह",
                "father": "मंटू साह",
                "mobile": "8809053570",
                "age": 28,
                "category": "OBC",
                "caste": "तेली",
                "profile": "जन सुराजी सामाजिक कार्यकर्ता",
                "calling_status": "",
                "remarks": "चुनाव अभियान समिति सदस्य",
                "struck": False
            },
            {
                "sno": 4,
                "block": "भवानीपुर",
                "panchayat": "",
                "role": "सदस्य",
                "name": "वरूण कु० सिंह",
                "father": "अवधेश श० सिंह",
                "mobile": "9123260168",
                "age": 45,
                "category": "General",
                "caste": "राजपूत",
                "profile": "वर्तमान सरपंच प्रतिनिधि",
                "calling_status": "",
                "remarks": "वर्तमान सरपंच प्रतिनिधि",
                "struck": False
            },
            {
                "sno": 5,
                "block": "भवानीपुर",
                "panchayat": "",
                "role": "सदस्य",
                "name": "अवधेश कु० सिंह",
                "father": "बलभद्र सिंह",
                "mobile": "9431869131",
                "age": 67,
                "category": "General",
                "caste": "राजपूत",
                "profile": "सेवानिवृत्त शिक्षक",
                "calling_status": "",
                "remarks": "सेवानिवृत्त शिक्षक",
                "struck": False
            },
            {
                "sno": 6,
                "block": "भवानीपुर",
                "panchayat": "",
                "role": "सदस्य",
                "name": "मनोज कु० भारती",
                "father": "स्व० फूलेश्वर प्र० यादव",
                "mobile": "6207797515",
                "age": 51,
                "category": "OBC",
                "caste": "यादव",
                "profile": "एमएलए प्रत्याशी रह चुके",
                "calling_status": "",
                "remarks": "पूर्व MLA प्रत्याशी",
                "struck": False
            },
            {
                "sno": 7,
                "block": "भवानीपुर",
                "panchayat": "",
                "role": "सदस्य",
                "name": "मो० लुकमान आलम",
                "father": "शेख भोला",
                "mobile": "9931731291",
                "age": 52,
                "category": "Minority",
                "caste": "मुस्लिम",
                "profile": "भूतपूर्व वार्ड सदस्य कई बार",
                "calling_status": "",
                "remarks": "कई बार पूर्व वार्ड सदस्य",
                "struck": False
            },
            {
                "sno": 8,
                "block": "भवानीपुर",
                "panchayat": "",
                "role": "सदस्य",
                "name": "अरविन्द कु० पासवान (उर्फ मुन्ना)",
                "father": "नरेश पासवान",
                "mobile": "6202834933",
                "age": 41,
                "category": "SC",
                "caste": "दुसाध",
                "profile": "वार्ड पार्षद कई बार",
                "calling_status": "",
                "remarks": "उर्फ मुन्ना | कई बार वार्ड पार्षद",
                "struck": False
            },
            {
                "sno": 9,
                "block": "भवानीपुर",
                "panchayat": "",
                "role": "सदस्य",
                "name": "सदानंद सहनी",
                "father": "विन्देश्वरी सहनी",
                "mobile": "8092391764",
                "age": 55,
                "category": "EBC",
                "caste": "गोढ़ी",
                "profile": "भूतपूर्व EBC जिला अध्यक्ष",
                "calling_status": "",
                "remarks": "पूर्व EBC जिलाध्यक्ष",
                "struck": False
            },
            {
                "sno": 10,
                "block": "भवानीपुर",
                "panchayat": "",
                "role": "सदस्य",
                "name": "शाहनवाज",
                "father": "स्व० नजीम",
                "mobile": "6200072457",
                "age": 30,
                "category": "Minority",
                "caste": "मुस्लिम",
                "profile": "नेता लोकल",
                "calling_status": "",
                "remarks": "स्थानीय नेता",
                "struck": False
            },
            {
                "sno": 11,
                "block": "भवानीपुर",
                "panchayat": "",
                "role": "सदस्य",
                "name": "विनय सिंह",
                "father": "स्व० सुरेन्द्र सिंह",
                "mobile": "9572513426",
                "age": 56,
                "category": "General",
                "caste": "राजपूत",
                "profile": "नेता लोकल",
                "calling_status": "",
                "remarks": "स्थानीय नेता",
                "struck": False
            },
            {
                "sno": 12,
                "block": "भवानीपुर",
                "panchayat": "",
                "role": "सदस्य",
                "name": "गुलाम रब्बानी",
                "father": "सोहराब आलम",
                "mobile": "6200129795",
                "age": 44,
                "category": "Minority",
                "caste": "मुस्लिम",
                "profile": "जन सुराजी नेता",
                "calling_status": "",
                "remarks": "जन सुराजी नेता",
                "struck": False
            },
            {
                "sno": 13,
                "block": "भवानीपुर",
                "panchayat": "",
                "role": "सदस्य",
                "name": "मुन्ना कु० यादव",
                "father": "दुख्खा प्रसाद यादव",
                "mobile": "9162819253",
                "age": 41,
                "category": "OBC",
                "caste": "यादव",
                "profile": "पूर्व वार्ड पार्षद",
                "calling_status": "",
                "remarks": "पूर्व वार्ड पार्षद",
                "struck": False
            }
        ]
    },
    {
        "sheet_title": "पूर्णियाँ पूर्व (Purnia East)",
        "block_name": "पूर्णियाँ पूर्व (Purnia East)",
        "district": "पूर्णियाँ (Purnia)",
        "source": "WhatsApp Image 2026-09-25 at 4.52.42 PM.jpeg",
        "records": [
            {
                "sno": 1,
                "block": "पूर्णियाँ पूर्व",
                "panchayat": "",
                "role": "प्रखंड प्रभारी",
                "name": "मो० इलियास",
                "father": "स्व० मो० इसहाक",
                "mobile": "9523034539",
                "age": "",
                "category": "Minority",
                "caste": "मुस्लिम",
                "profile": "प्रखंड प्रभारी",
                "calling_status": "",
                "remarks": "प्रखंड प्रभारी (हस्ताक्षरकर्ता)",
                "struck": False
            },
            {
                "sno": 2,
                "block": "पूर्णियाँ पूर्व",
                "panchayat": "",
                "role": "प्रखंड प्रभारी",
                "name": "अशोक यादव",
                "father": "स्व० हरिहर यादव",
                "mobile": "8051893758",
                "age": "",
                "category": "OBC",
                "caste": "यादव",
                "profile": "प्रखंड प्रभारी",
                "calling_status": "",
                "remarks": "प्रखंड प्रभारी (हस्ताक्षरकर्ता)",
                "struck": False
            },
            {
                "sno": 3,
                "block": "पूर्णियाँ पूर्व",
                "panchayat": "",
                "role": "प्रखंड चुनाव अभियान समिति सदस्य",
                "name": "गुलाम रसूल",
                "father": "",
                "mobile": "6201214418",
                "age": "",
                "category": "Minority",
                "caste": "मुस्लिम",
                "profile": "चुनाव अभियान समिति",
                "calling_status": "",
                "remarks": "अभियान समिति सदस्य",
                "struck": False
            },
            {
                "sno": 4,
                "block": "पूर्णियाँ पूर्व",
                "panchayat": "① बीरपुर / ② विक्रम पुर",
                "role": "सदस्य",
                "name": "मो० शमशाद आलम",
                "father": "स्व० अब्दुल कय्युम",
                "mobile": "9546246325",
                "age": "",
                "category": "Minority",
                "caste": "मुस्लिम",
                "profile": "",
                "calling_status": "",
                "remarks": "बीरपुर / विक्रम पुर",
                "struck": False
            },
            {
                "sno": 5,
                "block": "पूर्णियाँ पूर्व",
                "panchayat": "③ सिकन्दर पुर",
                "role": "सदस्य",
                "name": "इजहारूल हक",
                "father": "",
                "mobile": "8969937825",
                "age": "",
                "category": "Minority",
                "caste": "मुस्लिम",
                "profile": "",
                "calling_status": "",
                "remarks": "सिकन्दर पुर",
                "struck": False
            },
            {
                "sno": 6,
                "block": "पूर्णियाँ पूर्व",
                "panchayat": "④ डिमिया छत्रजान",
                "role": "सदस्य",
                "name": "मोईन",
                "father": "मो० रियाजउद्दीन",
                "mobile": "7353971244",
                "age": "",
                "category": "Minority",
                "caste": "मुस्लिम",
                "profile": "",
                "calling_status": "",
                "remarks": "डिमिया छत्रजान",
                "struck": False
            },
            {
                "sno": 7,
                "block": "पूर्णियाँ पूर्व",
                "panchayat": "⑤ रामपुर",
                "role": "सदस्य",
                "name": "अनुज जी",
                "father": "",
                "mobile": "8709214546",
                "age": "",
                "category": "",
                "caste": "",
                "profile": "",
                "calling_status": "",
                "remarks": "रामपुर",
                "struck": False
            },
            {
                "sno": 8,
                "block": "पूर्णियाँ पूर्व",
                "panchayat": "⑥ भेनाकरियात",
                "role": "सदस्य",
                "name": "अर्जुन मंडल",
                "father": "",
                "mobile": "9304568122",
                "age": "",
                "category": "OBC / EBC",
                "caste": "मंडल",
                "profile": "",
                "calling_status": "",
                "remarks": "भेनाकरियात",
                "struck": False
            },
            {
                "sno": 9,
                "block": "पूर्णियाँ पूर्व",
                "panchayat": "⑦ गौरा",
                "role": "सदस्य",
                "name": "संगम महतो",
                "father": "",
                "mobile": "7667938426",
                "age": "",
                "category": "OBC",
                "caste": "महतो",
                "profile": "",
                "calling_status": "",
                "remarks": "गौरा",
                "struck": False
            },
            {
                "sno": 10,
                "block": "पूर्णियाँ पूर्व",
                "panchayat": "⑧ बिथारपुर",
                "role": "सदस्य",
                "name": "हबीब",
                "father": "मो० घुसा",
                "mobile": "7631104965",
                "age": "",
                "category": "Minority",
                "caste": "मुस्लिम",
                "profile": "",
                "calling_status": "",
                "remarks": "बिथारपुर",
                "struck": False
            },
            {
                "sno": 11,
                "block": "पूर्णियाँ पूर्व",
                "panchayat": "⑨ कबैया",
                "role": "सदस्य",
                "name": "मो० शमीम अख्तर",
                "father": "फजलुर रहमान",
                "mobile": "8789842240",
                "age": "",
                "category": "Minority",
                "caste": "मुस्लिम",
                "profile": "",
                "calling_status": "",
                "remarks": "कबैया (लालगंज काटा गया)",
                "struck": False
            },
            {
                "sno": 12,
                "block": "पूर्णियाँ पूर्व",
                "panchayat": "⑪ हरदा",
                "role": "सदस्य",
                "name": "सदानन्द पासवान",
                "father": "ढोढई पासवान",
                "mobile": "8677065174",
                "age": "",
                "category": "SC",
                "caste": "पासवान",
                "profile": "",
                "calling_status": "",
                "remarks": "हरदा",
                "struck": False
            },
            {
                "sno": 13,
                "block": "पूर्णियाँ पूर्व",
                "panchayat": "⑫ चान्दी, राजीगंज",
                "role": "सदस्य",
                "name": "हेमन्त जी",
                "father": "",
                "mobile": "7004482572",
                "age": "",
                "category": "",
                "caste": "",
                "profile": "",
                "calling_status": "",
                "remarks": "चान्दी, राजीगंज",
                "struck": False
            },
            {
                "sno": 14,
                "block": "पूर्णियाँ पूर्व",
                "panchayat": "महाराजपुर",
                "role": "सदस्य",
                "name": "सचिन कु० सिंह",
                "father": "रामाश्रय सिंह",
                "mobile": "8210522031",
                "age": "",
                "category": "General",
                "caste": "राजपूत",
                "profile": "",
                "calling_status": "",
                "remarks": "महाराजपुर",
                "struck": False
            }
        ]
    },
    {
        "sheet_title": "बी० कोठी (Barhara Kothi)",
        "block_name": "बी० कोठी (Barhara Kothi)",
        "district": "पूर्णियाँ (Purnia)",
        "source": "WhatsApp Image 2026-09-25 at 4.52.42 PM (1).jpeg",
        "records": [
            {
                "sno": 1,
                "block": "बी० कोठी",
                "panchayat": "",
                "role": "MLC चुनाव समिति सदस्य",
                "name": "अरुण देव मंडल",
                "father": "",
                "mobile": "9199104834",
                "age": "",
                "category": "OBC / EBC",
                "caste": "मंडल",
                "profile": "MLC चुनाव हेतु गठित सदस्य",
                "calling_status": "",
                "remarks": "मो० 9199104834 / 9199107834",
                "struck": False
            },
            {
                "sno": 2,
                "block": "बी० कोठी",
                "panchayat": "",
                "role": "MLC चुनाव समिति सदस्य",
                "name": "गुलशन सिंह",
                "father": "",
                "mobile": "9955063916",
                "age": "",
                "category": "General",
                "caste": "राजपूत",
                "profile": "MLC चुनाव हेतु गठित सदस्य",
                "calling_status": "",
                "remarks": "",
                "struck": False
            },
            {
                "sno": 3,
                "block": "बी० कोठी",
                "panchayat": "",
                "role": "MLC चुनाव समिति सदस्य",
                "name": "आशुतोष कुमार चन्द",
                "father": "",
                "mobile": "9470048165",
                "age": "",
                "category": "",
                "caste": "चन्द",
                "profile": "MLC चुनाव हेतु गठित सदस्य",
                "calling_status": "",
                "remarks": "",
                "struck": False
            },
            {
                "sno": 4,
                "block": "बी० कोठी",
                "panchayat": "",
                "role": "MLC चुनाव समिति सदस्य",
                "name": "सोमन पासवान",
                "father": "",
                "mobile": "9631276225",
                "age": "",
                "category": "SC",
                "caste": "पासवान",
                "profile": "MLC चुनाव हेतु गठित सदस्य",
                "calling_status": "",
                "remarks": "",
                "struck": False
            },
            {
                "sno": 5,
                "block": "बी० कोठी",
                "panchayat": "",
                "role": "MLC चुनाव समिति सदस्य",
                "name": "खगेश राम",
                "father": "",
                "mobile": "9801535699",
                "age": "",
                "category": "SC",
                "caste": "राम (चमार)",
                "profile": "MLC चुनाव हेतु गठित सदस्य",
                "calling_status": "",
                "remarks": "",
                "struck": False
            },
            {
                "sno": 6,
                "block": "बी० कोठी",
                "panchayat": "",
                "role": "MLC चुनाव समिति सदस्य",
                "name": "आत्मानंद सागर",
                "father": "",
                "mobile": "9572777935",
                "age": "",
                "category": "",
                "caste": "",
                "profile": "MLC चुनाव हेतु गठित सदस्य",
                "calling_status": "",
                "remarks": "",
                "struck": False
            },
            {
                "sno": 7,
                "block": "बी० कोठी",
                "panchayat": "",
                "role": "MLC चुनाव समिति सदस्य",
                "name": "रोशन शर्मा",
                "father": "",
                "mobile": "6207043864",
                "age": "",
                "category": "EBC",
                "caste": "शर्मा (बढ़ई/लोहार)",
                "profile": "MLC चुनाव हेतु गठित सदस्य",
                "calling_status": "",
                "remarks": "",
                "struck": False
            },
            {
                "sno": 8,
                "block": "बी० कोठी",
                "panchayat": "",
                "role": "MLC चुनाव समिति सदस्य",
                "name": "खगेन्द्र शाह",
                "father": "",
                "mobile": "9507294729",
                "age": "",
                "category": "OBC",
                "caste": "शाह (तेली/साहू)",
                "profile": "MLC चुनाव हेतु गठित सदस्य",
                "calling_status": "",
                "remarks": "",
                "struck": False
            },
            {
                "sno": 9,
                "block": "बी० कोठी",
                "panchayat": "",
                "role": "MLC चुनाव समिति सदस्य",
                "name": "चन्दन झा",
                "father": "",
                "mobile": "8809215210",
                "age": "",
                "category": "General",
                "caste": "ब्राह्मण (झा)",
                "profile": "MLC चुनाव हेतु गठित सदस्य",
                "calling_status": "",
                "remarks": "",
                "struck": False
            },
            {
                "sno": 10,
                "block": "बी० कोठी",
                "panchayat": "",
                "role": "MLC चुनाव समिति सदस्य",
                "name": "गुलशन कुमार",
                "father": "",
                "mobile": "6200593135",
                "age": "",
                "category": "",
                "caste": "",
                "profile": "MLC चुनाव हेतु गठित सदस्य",
                "calling_status": "",
                "remarks": "",
                "struck": False
            },
            {
                "sno": 11,
                "block": "बी० कोठी",
                "panchayat": "",
                "role": "MLC चुनाव समिति सदस्य",
                "name": "सुभाष कुमार रजक",
                "father": "",
                "mobile": "8051777268",
                "age": "",
                "category": "SC",
                "caste": "धोबी (रजक)",
                "profile": "MLC चुनाव हेतु गठित सदस्य",
                "calling_status": "",
                "remarks": "",
                "struck": False
            }
        ]
    },
    {
        "sheet_title": "बनमनखी (Banmankhi)",
        "block_name": "बनमनखी (Banmankhi)",
        "district": "पूर्णियाँ (Purnia)",
        "source": "WhatsApp Image 2026-09-25 at 4.52.42 PM (2).jpeg",
        "records": [
            {
                "sno": 1,
                "block": "बनमनखी",
                "panchayat": "",
                "role": "प्रखंड प्रभारी",
                "name": "सागर कु० मिश्रा",
                "father": "प्रकाश मिश्रा",
                "mobile": "6200137322",
                "age": 32,
                "category": "General",
                "caste": "ब्राह्मण",
                "profile": "",
                "calling_status": "",
                "remarks": "प्रखंड प्रभारी",
                "struck": False
            },
            {
                "sno": 2,
                "block": "बनमनखी",
                "panchayat": "",
                "role": "प्रखंड प्रभारी",
                "name": "राजेन्द्र शर्मा",
                "father": "",
                "mobile": "",
                "age": "",
                "category": "EBC",
                "caste": "शर्मा",
                "profile": "",
                "calling_status": "",
                "remarks": "प्रखंड प्रभारी",
                "struck": False
            },
            {
                "sno": 3,
                "block": "बनमनखी",
                "panchayat": "",
                "role": "प्रखंड चुनाव अभियान समिति सदस्य",
                "name": "अजय भगत",
                "father": "",
                "mobile": "9939402741",
                "age": 40,
                "category": "OBC / EBC",
                "caste": "बनिया",
                "profile": "पंचायत चुनाव का अनुभव",
                "calling_status": "",
                "remarks": "पंचायत चुनाव का अनुभव",
                "struck": False
            },
            {
                "sno": 4,
                "block": "बनमनखी",
                "panchayat": "",
                "role": "सदस्य",
                "name": "मो० आरिफ",
                "father": "",
                "mobile": "9113479142",
                "age": 30,
                "category": "Minority",
                "caste": "मुसलमान",
                "profile": "",
                "calling_status": "",
                "remarks": "",
                "struck": False
            },
            {
                "sno": 5,
                "block": "बनमनखी",
                "panchayat": "",
                "role": "सदस्य",
                "name": "राकेश ऋषि",
                "father": "",
                "mobile": "9507109192",
                "age": 35,
                "category": "SC",
                "caste": "मुसहर",
                "profile": "",
                "calling_status": "",
                "remarks": "",
                "struck": False
            },
            {
                "sno": 6,
                "block": "बनमनखी",
                "panchayat": "",
                "role": "सदस्य",
                "name": "मानस यादव",
                "father": "",
                "mobile": "9931047580",
                "age": 45,
                "category": "OBC",
                "caste": "यादव",
                "profile": "वर्तमान मुखिया",
                "calling_status": "",
                "remarks": "वर्तमान मुखिया",
                "struck": False
            },
            {
                "sno": 7,
                "block": "बनमनखी",
                "panchayat": "",
                "role": "सदस्य",
                "name": "मनोज सिंह",
                "father": "",
                "mobile": "6207939417",
                "age": 45,
                "category": "General",
                "caste": "राजपूत",
                "profile": "",
                "calling_status": "",
                "remarks": "",
                "struck": False
            },
            {
                "sno": 8,
                "block": "बनमनखी",
                "panchayat": "",
                "role": "सदस्य",
                "name": "अंजली कु० दीपक",
                "father": "",
                "mobile": "7004482707",
                "age": 40,
                "category": "General",
                "caste": "ब्राह्मण",
                "profile": "",
                "calling_status": "",
                "remarks": "",
                "struck": False
            },
            {
                "sno": 9,
                "block": "बनमनखी",
                "panchayat": "",
                "role": "सदस्य",
                "name": "साजन चौधरी",
                "father": "",
                "mobile": "9570605525",
                "age": 35,
                "category": "OBC / EBC",
                "caste": "बनिया",
                "profile": "",
                "calling_status": "",
                "remarks": "",
                "struck": False
            },
            {
                "sno": 10,
                "block": "बनमनखी",
                "panchayat": "",
                "role": "सदस्य",
                "name": "देवानंद राम",
                "father": "",
                "mobile": "9771116097",
                "age": 50,
                "category": "SC",
                "caste": "हरिजन",
                "profile": "पूर्व मुखिया",
                "calling_status": "",
                "remarks": "पूर्व मुखिया (काटा गया शब्द)",
                "struck": False
            },
            {
                "sno": 11,
                "block": "बनमनखी",
                "panchayat": "",
                "role": "सदस्य",
                "name": "प्रवीण मंडल",
                "father": "",
                "mobile": "8298443303",
                "age": 42,
                "category": "OBC / EBC",
                "caste": "कुर्मी मंडल",
                "profile": "",
                "calling_status": "",
                "remarks": "",
                "struck": False
            },
            {
                "sno": 12,
                "block": "बनमनखी",
                "panchayat": "",
                "role": "सदस्य",
                "name": "विकाष आनंद",
                "father": "",
                "mobile": "9162997187",
                "age": 45,
                "category": "General",
                "caste": "ब्राह्मण",
                "profile": "",
                "calling_status": "",
                "remarks": "",
                "struck": False
            }
        ]
    },
    {
        "sheet_title": "कसबा (Kasba)",
        "block_name": "कसबा (Kasba)",
        "district": "पूर्णियाँ (Purnia)",
        "source": "WhatsApp Image 2026-09-25 at 4.52.43 PM.jpeg",
        "records": [
            {
                "sno": 1,
                "block": "कसबा",
                "panchayat": "बनैली (Banaili)",
                "role": "नेता / सदस्य",
                "name": "ध्रुव कु० झा",
                "father": "महेश्वरी झा",
                "mobile": "6299389557",
                "age": 39,
                "category": "General (Gen)",
                "caste": "Brahmin (ब्राह्मण)",
                "profile": "Three term Incumbent, Defacto ward member 2011, 2016 & 2021, Owns CSP",
                "calling_status": "Connected",
                "remarks": "3 बार वार्ड सदस्य, CSP संचालक",
                "struck": False
            },
            {
                "sno": 2,
                "block": "कसबा",
                "panchayat": "गुरही (Gurhi)",
                "role": "नेता / सदस्य",
                "name": "मोहम्मद अमीनउद्दीन",
                "father": "स्व० हाजी मुसहरू",
                "mobile": "9939885394",
                "age": 58,
                "category": "Minority",
                "caste": "Sheikhra (शेखरा)",
                "profile": "",
                "calling_status": "Switched off",
                "remarks": "कॉल स्थिति: स्विच ऑफ",
                "struck": False
            },
            {
                "sno": 3,
                "block": "कसबा",
                "panchayat": "बौचगांव (Bauchgaon)",
                "role": "नेता / सदस्य",
                "name": "मोहम्मद यूसुफ",
                "father": "मोहम्मद मोजिब",
                "mobile": "7765086673",
                "age": 32,
                "category": "Minority",
                "caste": "Kulhaiya (कुलहैया)",
                "profile": "",
                "calling_status": "Connected - Yes",
                "remarks": "कॉल स्थिति: कनेक्टेड (Yes)",
                "struck": False
            },
            {
                "sno": 4,
                "block": "कसबा",
                "panchayat": "कुल्ला खास (Kulla Khas)",
                "role": "नेता / सदस्य",
                "name": "रंजीत पासवान",
                "father": "फटकन पासवान",
                "mobile": "8709810750",
                "age": 45,
                "category": "SC",
                "caste": "Paswan (पासवान)",
                "profile": "",
                "calling_status": "Connected - Yes",
                "remarks": "कॉल स्थिति: कनेक्टेड (Yes)",
                "struck": False
            },
            {
                "sno": 5,
                "block": "कसबा",
                "panchayat": "सबदलपुर (Sabdalpur)",
                "role": "नेता / सदस्य",
                "name": "कैलाश महलदार",
                "father": "स्वर्गीय नगीना महलदार",
                "mobile": "7646046458",
                "age": 44,
                "category": "EBC",
                "caste": "Mallah (मल्लाह)",
                "profile": "",
                "calling_status": "Connected - Yes",
                "remarks": "कॉल स्थिति: कनेक्टेड (Yes)",
                "struck": False
            },
            {
                "sno": 6,
                "block": "कसबा",
                "panchayat": "मल्हारिया (Malhariya)",
                "role": "नेता / सदस्य",
                "name": "मोहम्मद इस्लामुद्दीन रज़ा",
                "father": "स्व० शेख सियादुर रहमान",
                "mobile": "9939648785",
                "age": 54,
                "category": "Minority",
                "caste": "Sheikhra (शेखरा)",
                "profile": "",
                "calling_status": "Connected - Unhealthy",
                "remarks": "कॉल स्थिति: कनेक्टेड (Unhealthy)",
                "struck": False
            },
            {
                "sno": 7,
                "block": "कसबा",
                "panchayat": "वार्ड नंबर 19 (Ward No. 19)",
                "role": "नेता / सदस्य",
                "name": "अजीत कु० साह",
                "father": "स्व० सीतम लाल साह",
                "mobile": "8581035031",
                "age": 58,
                "category": "OBC",
                "caste": "Kalwar (कलवार)",
                "profile": "",
                "calling_status": "Connected - Unhealthy",
                "remarks": "कॉल स्थिति: कनेक्टेड (Unhealthy)",
                "struck": False
            },
            {
                "sno": 8,
                "block": "कसबा",
                "panchayat": "मोधनी (Modhani)",
                "role": "नेता / सदस्य",
                "name": "गणेश सरकार",
                "father": "जियलाल सरकार",
                "mobile": "9142680313",
                "age": 64,
                "category": "OBC",
                "caste": "Kalwar (कलवार)",
                "profile": "",
                "calling_status": "Connected - Yes",
                "remarks": "कॉल स्थिति: कनेक्टेड (Yes)",
                "struck": False
            },
            {
                "sno": 9,
                "block": "कसबा",
                "panchayat": "अचली (Achli)",
                "role": "नेता / सदस्य",
                "name": "मतीन अंसारी",
                "father": "स्व० जलील अंसारी",
                "mobile": "6207394012",
                "age": 55,
                "category": "EBC",
                "caste": "Kewart (कैवर्त)",
                "profile": "",
                "calling_status": "Connected - Yes",
                "remarks": "कॉल स्थिति: कनेक्टेड (Yes)",
                "struck": False
            },
            {
                "sno": 10,
                "block": "कसबा",
                "panchayat": "वार्ड नं० 12 (Ward No. 12)",
                "role": "नेता / सदस्य",
                "name": "गोलटा ऋषि (सिंह)",
                "father": "राज कु० ऋ० (राज कुमार ऋषि)",
                "mobile": "9661442502",
                "age": 43,
                "category": "SC",
                "caste": "Musahar (मुसहर)",
                "profile": "",
                "calling_status": "",
                "remarks": "जाति: मुसहर",
                "struck": False
            },
            {
                "sno": 11,
                "block": "कसबा",
                "panchayat": "कुल्ला खास (Kulla Khas)",
                "role": "नेता / सदस्य",
                "name": "शिव शंकर सिंह",
                "father": "अम्बिका प्र० सिंह",
                "mobile": "9572116669",
                "age": 44,
                "category": "General",
                "caste": "Rajput (राजपूत)",
                "profile": "",
                "calling_status": "Connected - Yes",
                "remarks": "कॉल स्थिति: कनेक्टेड (Yes)",
                "struck": False
            },
            {
                "sno": 12,
                "block": "कसबा",
                "panchayat": "मझरा लागन (Majhra Lagan)",
                "role": "नेता / सदस्य",
                "name": "मो० हसन",
                "father": "स्व० मो० मोसीन",
                "mobile": "9931946537",
                "age": 55,
                "category": "Minority",
                "caste": "Kulhaiya (कुलहैया)",
                "profile": "",
                "calling_status": "Connected - Yes",
                "remarks": "कॉल स्थिति: कनेक्टेड (Yes)",
                "struck": False
            }
        ]
    }
]

def style_sheet(ws, title, subtitle, headers, records, is_master=False):
    ws.sheet_properties.tabColor = "1F4E79" if is_master else "2F75B6"
    
    # Title Banner
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    title_cell = ws.cell(row=1, column=1)
    title_cell.value = title
    title_cell.font = FONT_TITLE
    title_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 36
    
    # Subtitle / Info
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(headers))
    sub_cell = ws.cell(row=2, column=1)
    sub_cell.value = subtitle
    sub_cell.font = Font(name="Calibri", size=10, italic=True, color="595959")
    sub_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 20
    
    # Header Row 4
    ws.row_dimensions[4].height = 28
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col_idx, value=h)
        cell.font = FONT_HEADER
        cell.fill = FILL_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_THIN
        
    # Data Rows
    for row_idx, r in enumerate(records, start=5):
        ws.row_dimensions[row_idx].height = 22
        is_even = (row_idx % 2 == 0)
        default_fill = FILL_ZEBRA if is_even else FILL_WHITE
        fill_to_use = FILL_STRUCK if r.get("struck") else default_fill
        font_to_use = FONT_STRUCK if r.get("struck") else FONT_DATA
        
        row_values = [
            r["sno"],
            r.get("block", ""),
            r.get("panchayat", ""),
            r.get("role", ""),
            r.get("name", ""),
            r.get("father", ""),
            r.get("mobile", ""),
            r.get("age", ""),
            r.get("category", ""),
            r.get("caste", ""),
            r.get("profile", ""),
            r.get("calling_status", ""),
            r.get("remarks", ""),
            r.get("source", "")
        ]
        
        for col_idx, val in enumerate(row_values, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=str(val) if val != "" else "")
            cell.font = font_to_use
            cell.fill = fill_to_use
            cell.border = BORDER_THIN
            
            # Formatting & Alignment
            if col_idx in (1, 7, 8, 9, 12, 14):
                cell.alignment = ALIGN_CENTER
            elif col_idx in (2, 3, 4, 10):
                cell.alignment = ALIGN_CENTER
            else:
                cell.alignment = ALIGN_LEFT
                
    # Auto-fit column widths
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            if cell.row in (1, 2, 3):
                continue
            val_str = str(cell.value or '')
            char_len = len(val_str.encode('utf-8')) / 2 if any(ord(c) > 127 for c in val_str) else len(val_str)
            if char_len > max_len:
                max_len = char_len
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

# Build Workbook for Purnia Blocks
wb_purnia = openpyxl.Workbook()
if wb_purnia.active is not None:
    wb_purnia.remove(wb_purnia.active)

ws_purnia_master = wb_purnia.create_sheet(title="Master - पूर्णियाँ (Purnia)")
master_purnia_records = []
sno_p = 1
for b in purnia_blocks:
    for rec in b["records"]:
        m_rec = dict(rec)
        m_rec["sno"] = sno_p
        m_rec["block"] = b["block_name"]
        m_rec["source"] = b["source"]
        master_purnia_records.append(m_rec)
        sno_p += 1

style_sheet(
    ws_purnia_master,
    "📊 मास्टर डेटा शीट - पूर्णियाँ जिला सभी 6 प्रखंड (Purnia District Consolidated Master)",
    f"कुल प्रखंड: 6 | कुल रिकॉर्ड: {len(master_purnia_records)} | हस्तलिखित एवं मुद्रित प्रपत्रों से पूर्णतः डिजिटलकृत व सत्यापित",
    HEADERS,
    master_purnia_records,
    is_master=True
)

for b in purnia_blocks:
    safe_title = b["sheet_title"][:31]
    ws = wb_purnia.create_sheet(title=safe_title)
    records = []
    for rec in b["records"]:
        r = dict(rec)
        r["block"] = b["block_name"]
        r["source"] = b["source"]
        records.append(r)
    style_sheet(
        ws,
        f"📍 प्रखंड: {b['block_name']} (जिला: {b['district']})",
        f"स्रोत दस्तावेज़: {b['source']} | कुल सदस्य/प्रविष्टियाँ: {len(records)}",
        HEADERS,
        records,
        is_master=False
    )

purnia_out = r"c:\Users\alamn\Downloads\25 Sept\Digitalized_Purnia_Blocks.xlsx"
wb_purnia.save(purnia_out)
print(f"Purnia workbook saved to {purnia_out} with {len(wb_purnia.sheetnames)} sheets.")


# 14 Original Blocks from Nalanda & Nawada Districts
blocks_data = [
    {
        "sheet_title": "अस्थावाँ (Asthawan)",
        "block_name": "अस्थावाँ (Asthawan)",
        "district": "Nalanda",
        "source": "IMG-20260925-WA0012.jpg",
        "records": [
            {"sno": 1, "name": "शशीकान्त जी", "role": "", "mobile": "7004771314", "location": "", "father": "", "remarks": "सक्रिय", "struck": False},
            {"sno": 2, "name": "पप्पू जी", "role": "प्रखंड अध्यक्ष", "mobile": "7091951443", "location": "", "father": "", "remarks": "प्रखंड अध्यक्ष (अंडरलाइन/चिह्नित)", "struck": False},
            {"sno": 3, "name": "पूनम कुमारी", "role": "पूर्व प्रखंड अध्यक्ष", "mobile": "8252968508", "location": "", "father": "", "remarks": "काटा गया (Crossed out)", "struck": True},
            {"sno": 4, "name": "त्रिपुरारी / विपुसती प्रसाद", "role": "", "mobile": "9155397826", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 5, "name": "Anjani (अंजनी)", "role": "Professor (प्रोफेसर)", "mobile": "9905821301", "location": "Bihar Sharif (बिहार शरीफ)", "father": "", "remarks": "", "struck": False},
            {"sno": 6, "name": "सीताराम शर्मा", "role": "", "mobile": "8709854480", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 7, "name": "नवीन कुमार", "role": "", "mobile": "7991142768", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 8, "name": "रीना देवी", "role": "", "mobile": "7070382929", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 9, "name": "राजीव यादव", "role": "", "mobile": "7992443566", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 10, "name": "विजित कुमार सिंह", "role": "", "mobile": "9934296454", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 11, "name": "शिवेन्द्र कुमार", "role": "", "mobile": "8237961017", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 12, "name": "शत्रुघ्न कुमार", "role": "", "mobile": "8434760207", "location": "", "father": "", "remarks": "पेज के शीर्ष पर अंकित संपर्क", "struck": False},
        ]
    },
    {
        "sheet_title": "इस्लामपुर (Islampur)",
        "block_name": "इस्लामपुर (Islampur)",
        "district": "Nalanda",
        "source": "IMG-20260925-WA0013.jpg",
        "records": [
            {"sno": 1, "name": "Kundan Kumar (कुंदन कुमार)", "role": "जिला युवा सचिव", "mobile": "9955749971", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 2, "name": "विशाल सिंह", "role": "नगर पंचायत युवा अध्यक्ष", "mobile": "8579852370", "location": "नगर पंचायत", "father": "", "remarks": "", "struck": False},
            {"sno": 3, "name": "स्वराज शर्मा", "role": "", "mobile": "6206013990", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 4, "name": "रितिक रौशन सिंह", "role": "ग्रामीण युवा अध्यक्ष", "mobile": "8709976284", "location": "ग्रामीण", "father": "", "remarks": "अंडरलाइन / काटा गया", "struck": False},
            {"sno": 5, "name": "पप्पू माहतो", "role": "", "mobile": "9693926168", "location": "रतनपुरा", "father": "", "remarks": "", "struck": False},
            {"sno": 6, "name": "पंकज शर्मा", "role": "", "mobile": "9899181438", "location": "परसुराय", "father": "", "remarks": "", "struck": False},
            {"sno": 7, "name": "राजू यादव", "role": "", "mobile": "7859021606", "location": "बेलदार विगहा", "father": "", "remarks": "", "struck": False},
            {"sno": 8, "name": "राजेश कुमार", "role": "", "mobile": "6202905372", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 9, "name": "रीता देवी", "role": "", "mobile": "9771674217", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 10, "name": "सोभी राम", "role": "", "mobile": "9510664819 / 9570664819", "location": "", "father": "", "remarks": "नंबर 9510 या 9570", "struck": False},
            {"sno": 11, "name": "राज शेखर", "role": "", "mobile": "9229752773", "location": "", "father": "", "remarks": "", "struck": False},
        ]
    },
    {
        "sheet_title": "नवादा-MLC (Nawada)",
        "block_name": "नवादा / वारिसलीगंज (Nawada / Warisaliganj)",
        "district": "Nawada",
        "source": "IMG-20260925-WA0014.jpg",
        "records": [
            {"sno": 1, "name": "Triven Singh (त्रिवेणी सिंह)", "role": "", "mobile": "", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 2, "name": "Sonu Prasad Kushwaha (सोनू प्रसाद कुशवाहा)", "role": "", "mobile": "", "location": "", "father": "", "remarks": "Checked (✓)", "struck": False},
            {"sno": 3, "name": "Amit Kumar (अमित कुमार)", "role": "", "mobile": "9113309098", "location": "", "father": "", "remarks": "काटा गया (Crossed out)", "struck": True},
            {"sno": 4, "name": "Subodh Singh (सुबोध सिंह)", "role": "", "mobile": "9934016823", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 5, "name": "Lalan Paswan (ललन पासवान)", "role": "", "mobile": "", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 6, "name": "Poonam Devi (पूनम देवी)", "role": "", "mobile": "", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 7, "name": "Niraj Kumar (नीरज कुमार)", "role": "", "mobile": "", "location": "", "father": "", "remarks": "Checked (✓)", "struck": False},
            {"sno": 8, "name": "Jogeshwar Verma / Jagdishwar Barma (जोगेश्वर वर्मा)", "role": "", "mobile": "", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 9, "name": "Sweety Devi (स्वीटी देवी)", "role": "", "mobile": "", "location": "", "father": "", "remarks": "Circled / xxxx marked", "struck": False},
            {"sno": 10, "name": "Naresh Chaudhary (नरेश चौधरी)", "role": "", "mobile": "", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 11, "name": "Hazi Zubair (हाजी जुबैर)", "role": "", "mobile": "", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 12, "name": "Mukesh Kumar Bablu (मुकेश कुमार बबलू)", "role": "", "mobile": "", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 13, "name": "Rajnesh Kumar (रजनीश कुमार)", "role": "", "mobile": "9431461219", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 14, "name": "Pt. Ajay Kumar (पं. अजय कुमार)", "role": "", "mobile": "", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 15, "name": "Rahul Raj (राहुल राज)", "role": "वारिसलीगंज प्रखंड अध्यक्ष (Warisaliganj Prakhand Adhyaksh)", "mobile": "8873087284", "location": "Panchayat Nepura (नेपुरा), Warisaliganj, Nawada", "father": "Binay Kumar (विनय कुमार)", "remarks": "विशेष बॉक्स नोट (Need to add)", "struck": False},
        ]
    },
    {
        "sheet_title": "कतरीसराय (Katrisarai)",
        "block_name": "कतरीसराय (Katrisarai)",
        "district": "Nalanda",
        "source": "IMG-20260925-WA0015.jpg",
        "records": [
            {"sno": 1, "name": "विनोद सिंह", "role": "प्रखंड अध्यक्ष", "mobile": "9534964609", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 2, "name": "बिहारी उर्फ भरत भूषण प्रसाद", "role": "", "mobile": "8235816516", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 3, "name": "परमानन्द सिंह", "role": "", "mobile": "7070091638", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 4, "name": "धर्मवीर कुमार", "role": "महामंत्री", "mobile": "8292420726", "location": "", "father": "", "remarks": "रेखांकित", "struck": False},
            {"sno": 5, "name": "पिंकी कुमारी कुशवाहा", "role": "महिला अध्यक्ष", "mobile": "6200575791", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 6, "name": "सुशील कुमार उर्फ बब्लू सिंह", "role": "", "mobile": "9308563910", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 7, "name": "अरविन्द प्रसाद सिंह", "role": "", "mobile": "9931752670", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 8, "name": "मिथिलेश पासवान", "role": "", "mobile": "8292471334", "location": "", "father": "", "remarks": "काटा गया (Crossed out)", "struck": True},
            {"sno": 9, "name": "बिरेन्द्र मांझी", "role": "", "mobile": "9934996756", "location": "", "father": "", "remarks": "मिथिलेश पासवान की जगह दर्ज", "struck": False},
            {"sno": 10, "name": "पूजा कुमारी", "role": "", "mobile": "6203866678 / 6203866648", "location": "", "father": "", "remarks": "अंतिम अंक 78 / 48", "struck": False},
            {"sno": 11, "name": "मुकुन्द कुमार", "role": "", "mobile": "", "location": "", "father": "", "remarks": "काटा गया (Crossed out)", "struck": True},
            {"sno": 12, "name": "नीतिश कुमार", "role": "", "mobile": "7667646461", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 13, "name": "मनोज कुमार", "role": "", "mobile": "", "location": "कताडीह", "father": "", "remarks": "", "struck": False},
            {"sno": 14, "name": "रविन्द्र सिंह", "role": "", "mobile": "9631513567", "location": "", "father": "", "remarks": "", "struck": False},
        ]
    },
    {
        "sheet_title": "नूरसराय (Noorsarai)",
        "block_name": "नूरसराय (Noorsarai)",
        "district": "Nalanda",
        "source": "IMG-20260925-WA0016.jpg",
        "records": [
            {"sno": 1, "name": "राजू कुमार", "role": "युवा सचिव", "mobile": "8540956103", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 2, "name": "कृष्णमुरारी यादव", "role": "प्रखंड उपाध्यक्ष", "mobile": "6204070480", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 3, "name": "बिपिन बिहारी", "role": "", "mobile": "9631278820", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 4, "name": "अवधेश रमण जी", "role": "जिला उपाध्यक्ष", "mobile": "8789697718", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 5, "name": "भोला कुमार", "role": "प्रखंड युवा अध्यक्ष", "mobile": "7061291358", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 6, "name": "बरूण कुमार", "role": "वार्ड सदस्य", "mobile": "9263405040", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 7, "name": "रजनी कुमार", "role": "", "mobile": "6204449807 / 6204149807", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 8, "name": "नवीन सिंह", "role": "पूर्व पंचायत समिति सदस्य", "mobile": "6207615579", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 9, "name": "राजीव कुमार सिंह", "role": "", "mobile": "7250001400", "location": "गजराज विगहा", "father": "", "remarks": "", "struck": False},
            {"sno": 10, "name": "विपिन कुमार", "role": "", "mobile": "9631278820", "location": "कपरसरि", "father": "", "remarks": "काटा गया / पुनः अंकित", "struck": True},
            {"sno": 11, "name": "गणेश कुमार", "role": "", "mobile": "9142259143", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 12, "name": "काजल कुमारी", "role": "", "mobile": "7645908029", "location": "", "father": "", "remarks": "", "struck": False},
        ]
    },
    {
        "sheet_title": "हरनौत (Harnaut)",
        "block_name": "हरनौत (Harnaut)",
        "district": "Nalanda",
        "source": "IMG-20260925-WA0017.jpg",
        "records": [
            {"sno": 1, "name": "मुन्ना जी", "role": "पूर्व महासचिव", "mobile": "8235801918", "location": "ढोंगरा", "father": "", "remarks": "", "struck": False},
            {"sno": 2, "name": "राजेश कुमार शर्मा", "role": "", "mobile": "8408923236", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 3, "name": "सोनी सिन्हा", "role": "", "mobile": "9525144204", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 4, "name": "गौतम जी", "role": "पूर्व प्रखंड अध्यक्ष", "mobile": "8825373801", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 5, "name": "अनुज गिरि", "role": "", "mobile": "7549978390", "location": "छतियाना", "father": "", "remarks": "", "struck": False},
            {"sno": 6, "name": "भास्कर कुमार", "role": "", "mobile": "9113449521", "location": "छतियाना", "father": "", "remarks": "", "struck": False},
            {"sno": 7, "name": "राहुल रंजन सिन्हा", "role": "", "mobile": "9304070094", "location": "बस्ती", "father": "", "remarks": "", "struck": False},
            {"sno": 8, "name": "शत्रुघ्न कुमार", "role": "", "mobile": "6203501947", "location": "वराह (बराह)", "father": "रामराज यादव", "remarks": "", "struck": False},
            {"sno": 9, "name": "Balmiki Kumar (वाल्मीकी कुमार)", "role": "", "mobile": "7079894063", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 10, "name": "वाल्मीकी पासवान", "role": "", "mobile": "8051717807", "location": "", "father": "", "remarks": "काटा गया (Crossed out)", "struck": True},
            {"sno": 11, "name": "डाॅ. सुभाष कुमार", "role": "", "mobile": "7004628803", "location": "हसनपुर धोसगांवा (वार्ड)", "father": "", "remarks": "", "struck": False},
        ]
    },
    {
        "sheet_title": "परबलपुर (Parbalpur)",
        "block_name": "परबलपुर (Parbalpur)",
        "district": "Nalanda",
        "source": "IMG-20260925-WA0018.jpg",
        "records": [
            {"sno": 1, "name": "खुशदिल कुमार", "role": "", "mobile": "9931357464", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 2, "name": "नितीश कुमार", "role": "प्रखंड अध्यक्ष", "mobile": "8409379610 / 8409307961", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 3, "name": "कुणाल कुमार", "role": "", "mobile": "9006021388", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 4, "name": "अंजनी कुमार", "role": "", "mobile": "7549414206", "location": "", "father": "", "remarks": "काटा गया (Crossed out)", "struck": True},
            {"sno": 5, "name": "रौशन सिंह", "role": "युवा अध्यक्ष", "mobile": "6203083549", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 6, "name": "अरविन्द कुमार शर्मा", "role": "", "mobile": "8815099088", "location": "", "father": "", "remarks": "काटा गया (Crossed out)", "struck": True},
            {"sno": 7, "name": "वसंत सिंह", "role": "", "mobile": "9771674698", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 8, "name": "प्रिंस कुमार", "role": "", "mobile": "7909083118", "location": "शंकरडीह", "father": "उदय कुमार", "remarks": "उम्र: 32 वर्ष", "struck": False},
            {"sno": 9, "name": "संजीव कुमार", "role": "", "mobile": "9570429434", "location": "धतारी विगहा", "father": "स्व० श्यामकिशोर सिंह", "remarks": "उम्र: 37 वर्ष", "struck": False},
            {"sno": 10, "name": "कन्हैया (रामनिवास कुमार)", "role": "", "mobile": "8434329221", "location": "दरियापुर", "father": "मिथिलेश कुमार", "remarks": "काटा गया (Crossed out)", "struck": True},
            {"sno": 11, "name": "कुन्दन कुमार", "role": "", "mobile": "6204170352", "location": "", "father": "संजय सिंह", "remarks": "", "struck": False},
            {"sno": 12, "name": "चन्दन कुमार", "role": "", "mobile": "7463801063", "location": "धोतीपुर", "father": "विजय सिंह", "remarks": "", "struck": False},
        ]
    },
    {
        "sheet_title": "चंडी (Chandi)",
        "block_name": "चंडी (Chandi)",
        "district": "Nalanda",
        "source": "IMG-20260925-WA0019.jpg",
        "records": [
            {"sno": 1, "name": "नदीम पासवान", "role": "जिला युवा सचिव", "mobile": "6207237242 / 8804990354", "location": "", "father": "", "remarks": "दो मोबाइल नंबर दर्ज", "struck": False},
            {"sno": 2, "name": "भूषण पासवान", "role": "", "mobile": "", "location": "", "father": "", "remarks": "पारस पासवान काटा गया", "struck": False},
            {"sno": 3, "name": "राजेश कुमार", "role": "", "mobile": "7488697406", "location": "बढ़ा", "father": "", "remarks": "", "struck": False},
            {"sno": 4, "name": "निर्मला कुमारी", "role": "", "mobile": "9341300520", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 5, "name": "आलोक कुमार", "role": "", "mobile": "6201813426", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 6, "name": "रूपेश कुमार यादव", "role": "", "mobile": "6301698100", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 7, "name": "नीतीश कुमार", "role": "", "mobile": "6299830521", "location": "", "father": "अभयकांत सिंह", "remarks": "", "struck": False},
            {"sno": 8, "name": "बच्चा पाठक", "role": "", "mobile": "7070371628", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 9, "name": "निर्भय कुमार चन्द्रवंशी", "role": "", "mobile": "7358922306", "location": "", "father": "", "remarks": "काटा गया (Crossed out)", "struck": True},
            {"sno": 10, "name": "पिन्टू रविदास", "role": "", "mobile": "9199772766", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 11, "name": "राकेश पासवान", "role": "", "mobile": "6206652272", "location": "", "father": "", "remarks": "दाएं हाशिए पर लंबवत अंकित", "struck": False},
        ]
    },
    {
        "sheet_title": "बिहार शरीफ (Bihar Sharif)",
        "block_name": "बिहार शरीफ (Bihar Sharif)",
        "district": "Nalanda",
        "source": "IMG-20260925-WA0020.jpg",
        "records": [
            {"sno": 1, "name": "सौरभ कुमार", "role": "", "mobile": "7739360070", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 2, "name": "गौरव कुमार", "role": "", "mobile": "8709192528", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 3, "name": "रॉकी कुमार", "role": "", "mobile": "8969570093", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 4, "name": "संजय सिंह", "role": "", "mobile": "6209730065", "location": "पावा", "father": "", "remarks": "अजय सिंह काटा गया", "struck": False},
            {"sno": 5, "name": "Satish Kumar (सतीश कुमार)", "role": "", "mobile": "9334021639", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 6, "name": "प्रेम सागर", "role": "", "mobile": "7667736307", "location": "तेतरावाँ", "father": "", "remarks": "", "struck": False},
            {"sno": 7, "name": "रचित कुमार", "role": "", "mobile": "7858988553", "location": "तुँगी", "father": "", "remarks": "", "struck": False},
            {"sno": 8, "name": "छोटु कुमार", "role": "", "mobile": "9122103016", "location": "मुरौरा", "father": "", "remarks": "", "struck": False},
            {"sno": 9, "name": "माईकल", "role": "", "mobile": "8271877141", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 10, "name": "संजय कुमार", "role": "", "mobile": "9835062707", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 11, "name": "राहुल", "role": "", "mobile": "7870170951", "location": "", "father": "अनील कुमार", "remarks": "रॉबिन कुमार काटा गया", "struck": False},
            {"sno": 12, "name": "सुबोध कुमार", "role": "", "mobile": "6200572626", "location": "", "father": "स्व० विशुन साव", "remarks": "", "struck": False},
            {"sno": 13, "name": "बब्लू कुमार", "role": "", "mobile": "8102567087", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 14, "name": "युवराज कुमार उर्फ सन्नी", "role": "", "mobile": "9122089472", "location": "", "father": "", "remarks": "", "struck": False},
        ]
    },
    {
        "sheet_title": "एकंगरसराय (Ekangarsarai)",
        "block_name": "एकंगरसराय / एकंगर (Ekangarsarai)",
        "district": "Nalanda",
        "source": "IMG-20260925-WA0021.jpg",
        "records": [
            {"sno": 1, "name": "मनोज कुमार", "role": "पूर्व प्रखंड अध्यक्ष", "mobile": "6201465903", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 2, "name": "सुन्दर यादव", "role": "", "mobile": "6205516132", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 3, "name": "नरेश यादव", "role": "", "mobile": "6201230327", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 4, "name": "विश्वकर्मा केसरी", "role": "", "mobile": "9122996075", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 5, "name": "कुमुद विश्वकर्मा", "role": "", "mobile": "9801221266", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 6, "name": "श्री देवी", "role": "", "mobile": "7631621747", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 7, "name": "प्रमोद कुमार प्रवीण", "role": "", "mobile": "9508278501", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 8, "name": "राकेश कुमार", "role": "", "mobile": "8541861483", "location": "", "father": "", "remarks": "", "struck": False},
        ]
    },
    {
        "sheet_title": "कराय परसुराय (Karai Parsurai)",
        "block_name": "कराय परसुराय (Karai Parsurai)",
        "district": "Nalanda",
        "source": "IMG-20260925-WA0022.jpg",
        "records": [
            {"sno": 1, "name": "अखिलेश पासवान", "role": "", "mobile": "8797618937", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 2, "name": "विजेन्द्र बिन्द", "role": "", "mobile": "9973782204", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 3, "name": "श्याम किशोर जी", "role": "", "mobile": "7870271930", "location": "", "father": "", "remarks": "", "struck": False},
        ]
    },
    {
        "sheet_title": "सिलाव (Silao)",
        "block_name": "सिलाव (Silao)",
        "district": "Nalanda",
        "source": "IMG-20260925-WA0023.jpg / WA0024.jpg",
        "records": [
            {"sno": 1, "name": "सुमन कुमार गुंजन", "role": "", "mobile": "9386561179", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 2, "name": "मुकुन्द शर्मा", "role": "", "mobile": "8603550935", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 3, "name": "दीपक कुमार", "role": "", "mobile": "7004286934", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 4, "name": "Chamcham Kumar (छमछम कुमार)", "role": "", "mobile": "8709440548", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 5, "name": "Jitendra Kumar (जितेन्द्र कुमार)", "role": "", "mobile": "8809496457", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 6, "name": "पिंकी कुमारी", "role": "", "mobile": "9193048587 / 9304858700", "location": "", "father": "", "remarks": "91-93048587..", "struck": False},
            {"sno": 7, "name": "राजू कुमार पाण्डेय", "role": "", "mobile": "8507732407", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 8, "name": "शेखर सिंह", "role": "", "mobile": "9955973622", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 9, "name": "सूरज कुमार", "role": "", "mobile": "9939646672", "location": "", "father": "", "remarks": "हाशिए पर अंकित", "struck": False},
        ]
    },
    {
        "sheet_title": "राजगीर (Rajgir)",
        "block_name": "राजगीर (Rajgir)",
        "district": "Nalanda",
        "source": "IMG-20260925-WA0024.jpg",
        "records": [
            {"sno": 1, "name": "शत्रुघ्न कुमार कुशवाहा", "role": "प्रखंड युवा अध्यक्ष", "mobile": "8434760207", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 2, "name": "अभय सिंह", "role": "", "mobile": "6202797006", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 3, "name": "सुनील कुमार", "role": "पूर्व अनुमंडल अध्यक्ष", "mobile": "9608164436", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 4, "name": "चुनचुन सिंह", "role": "प्रवक्ता", "mobile": "7209318832", "location": "नई पोखर", "father": "", "remarks": "", "struck": False},
            {"sno": 5, "name": "सरोज देवी", "role": "", "mobile": "7371084174", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 6, "name": "अनील कुमार", "role": "", "mobile": "", "location": "", "father": "", "remarks": "", "struck": False},
        ]
    },
    {
        "sheet_title": "थरथरी (Tharthari)",
        "block_name": "थरथरी (Tharthari)",
        "district": "Nalanda",
        "source": "IMG-20260925-WA0025.jpg",
        "records": [
            {"sno": 1, "name": "अरूण सिंह", "role": "", "mobile": "7319628742", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 2, "name": "सुबोध सिंह", "role": "", "mobile": "9939227462", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 3, "name": "अर्चना पाण्डेय", "role": "", "mobile": "7982574928", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 4, "name": "तारा पाण्डेय", "role": "", "mobile": "8084435237", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 5, "name": "उमाशंकर चौहान", "role": "", "mobile": "9576761251", "location": "", "father": "", "remarks": "", "struck": False},
            {"sno": 6, "name": "हिमांशु कुमार", "role": "", "mobile": "7979003950", "location": "", "father": "रामानुज कुमार", "remarks": "", "struck": False},
            {"sno": 7, "name": "दयानन्द", "role": "", "mobile": "", "location": "", "father": "", "remarks": "", "struck": False},
        ]
    },
]

old_blocks_data = blocks_data
# Build Combined 20 Blocks Workbook
wb_all = openpyxl.Workbook()
if wb_all.active is not None:
    wb_all.remove(wb_all.active)

# Master Sheet for All 20 Blocks
ws_all_master = wb_all.create_sheet(title="Master - सभी 20 प्रखंड")
all_master_records = []
sno_all = 1

# Add Nalanda/Nawada records
for b in old_blocks_data:
    for rec in b["records"]:
        m_rec = {
            "sno": sno_all,
            "block": b["block_name"],
            "panchayat": rec.get("location", ""),
            "role": rec.get("role", ""),
            "name": rec.get("name", ""),
            "father": rec.get("father", ""),
            "mobile": rec.get("mobile", ""),
            "age": "",
            "category": "",
            "caste": "",
            "profile": "",
            "calling_status": "",
            "remarks": rec.get("remarks", ""),
            "source": b.get("source", ""),
            "struck": rec.get("struck", False)
        }
        all_master_records.append(m_rec)
        sno_all += 1

# Add Purnia records
for b in purnia_blocks:
    for rec in b["records"]:
        m_rec = dict(rec)
        m_rec["sno"] = sno_all
        m_rec["block"] = b["block_name"]
        m_rec["source"] = b["source"]
        all_master_records.append(m_rec)
        sno_all += 1

style_sheet(
    ws_all_master,
    "📊 संपूर्ण मास्टर डेटा शीट - नालंदा, नवादा एवं पूर्णियाँ (20 प्रखंड Consolidated Master)",
    f"कुल प्रखंड: 20 | कुल प्रविष्टियाँ: {len(all_master_records)} | पूर्णतः सत्यापित एवं डिजिटलकृत",
    HEADERS,
    all_master_records,
    is_master=True
)

# Add sheets for Nalanda/Nawada blocks
for b in old_blocks_data:
    raw_title = str(b.get("sheet_title", ""))
    safe_title = raw_title.replace("/", "-").replace(":", "")[:31]
    ws = wb_all.create_sheet(title=safe_title)
    b_records = []
    for rec in b["records"]:
        r = {
            "sno": rec["sno"],
            "block": b["block_name"],
            "panchayat": rec.get("location", ""),
            "role": rec.get("role", ""),
            "name": rec.get("name", ""),
            "father": rec.get("father", ""),
            "mobile": rec.get("mobile", ""),
            "age": "",
            "category": "",
            "caste": "",
            "profile": "",
            "calling_status": "",
            "remarks": rec.get("remarks", ""),
            "source": b.get("source", ""),
            "struck": rec.get("struck", False)
        }
        b_records.append(r)
    style_sheet(
        ws,
        f"📍 प्रखंड: {b['block_name']} (जिला: {b['district']})",
        f"स्रोत दस्तावेज़: {b['source']} | कुल सदस्य: {len(b_records)}",
        HEADERS,
        b_records,
        is_master=False
    )

# Add sheets for Purnia blocks
for b in purnia_blocks:
    safe_title = b["sheet_title"][:31]
    ws = wb_all.create_sheet(title=safe_title)
    records = []
    for rec in b["records"]:
        r = dict(rec)
        r["block"] = b["block_name"]
        r["source"] = b["source"]
        records.append(r)
    style_sheet(
        ws,
        f"📍 प्रखंड: {b['block_name']} (जिला: {b['district']})",
        f"स्रोत दस्तावेज़: {b['source']} | कुल सदस्य/प्रविष्टियाँ: {len(records)}",
        HEADERS,
        records,
        is_master=False
    )

all_out = r"c:\Users\alamn\Downloads\25 Sept\Digitalized_Block_Data.xlsx"
wb_all.save(all_out)
print(f"Complete 20-block workbook saved to {all_out} with {len(wb_all.sheetnames)} sheets.")
