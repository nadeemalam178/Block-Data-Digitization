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

new_blocks_data = [
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
    
    # Empty Row 3
    ws.row_dimensions[3].height = 10
    
    # Header Row 4
    ws.row_dimensions[4].height = 28
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col_idx, value=h)
        cell.font = FONT_HEADER
        cell.fill = FILL_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_THIN
        
    # Data Rows starting at Row 5
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
            
            # Alignments
            # S.No (1), Mobile (7), Age (8), Category (9), Calling Status (12), Source (14)
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

def generate_purnia_workbook():
    wb = openpyxl.Workbook()
    if wb.active is not None:
        wb.remove(wb.active)
        
    # 1. Master Tab for Purnia Blocks
    ws_master = wb.create_sheet(title="Master - सभी प्रखंड (All Blocks)")
    master_records = []
    master_sno = 1
    for b in new_blocks_data:
        for rec in b["records"]:
            m_rec = dict(rec)
            m_rec["sno"] = master_sno
            m_rec["block"] = b["block_name"]
            m_rec["source"] = b["source"]
            master_records.append(m_rec)
            master_sno += 1
            
    style_sheet(
        ws_master,
        "📊 मास्टर डेटा शीट - पूर्णियाँ जिला सभी प्रखंड (Consolidated Purnia District Master List)",
        f"कुल प्रखंड: 6 | कुल रिकॉर्ड: {len(master_records)} | हस्तलिखित एवं मुद्रित प्रपत्रों से पूर्णतः डिजिटलकृत व सत्यापित",
        HEADERS,
        master_records,
        is_master=True
    )
    
    # 2. Individual Block Tabs
    for b in new_blocks_data:
        raw_title = b["sheet_title"]
        safe_title = raw_title.replace("/", "-").replace(":", "")[:31]
        ws = wb.create_sheet(title=safe_title)
        
        block_records = []
        for rec in b["records"]:
            b_rec = dict(rec)
            b_rec["block"] = b["block_name"]
            b_rec["source"] = b["source"]
            block_records.append(b_rec)
            
        style_sheet(
            ws,
            f"📍 प्रखंड: {b['block_name']} (जिला: {b['district']})",
            f"स्रोत दस्तावेज़: {b['source']} | कुल सदस्य/प्रविष्टियाँ: {len(block_records)}",
            HEADERS,
            block_records,
            is_master=False
        )
        
    out_file = r"c:\Users\alamn\Downloads\25 Sept\Digitalized_Purnia_Blocks.xlsx"
    wb.save(out_file)
    print(f"Purnia workbook saved successfully: {out_file} with {len(wb.sheetnames)} sheets.")

if __name__ == "__main__":
    generate_purnia_workbook()
