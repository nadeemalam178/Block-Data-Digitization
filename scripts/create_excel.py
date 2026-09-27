import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Create workbook
wb = openpyxl.Workbook()
# Remove default sheet
if wb.active is not None:
    wb.remove(wb.active)

# Styles
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
    "प्रखंड / क्षेत्र (Block / Area)",
    "नाम (Name)",
    "पद / दायित्व (Designation / Role)",
    "मोबाइल नंबर (Mobile No.)",
    "स्थान / ग्राम / पंचायत (Location / Village)",
    "पिता / अभिभावक का नाम (Father / Guardian)",
    "स्थिति / टिप्पणी (Status / Remarks)",
    "स्रोत चित्र (Source Image)"
]

# Data definition for all 14 blocks
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

def style_sheet(ws, title, headers, records, is_master=False):
    # Set tab color
    ws.sheet_properties.tabColor = "1F4E79" if is_master else "2F75B6"
    
    # Title Banner
    ws.merge_cells("A1:I1")
    title_cell = ws["A1"]
    title_cell.value = title
    title_cell.font = FONT_TITLE
    title_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 36
    
    # Subtitle / Info
    ws.merge_cells("A2:I2")
    sub_cell = ws["A2"]
    sub_cell.value = "डिजिटाइज़्ड ब्लॉक डेटा (हस्तलिखित प्रविष्टियों से संकलित)" if is_master else f"हस्तलिखित डेटा डिजिटलीकरण - स्रोत: {records[0]['source'] if records and 'source' in records[0] else ''}"
    sub_cell.font = Font(name="Calibri", size=10, italic=True, color="595959")
    sub_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 20
    
    # Header Row
    ws.row_dimensions[4].height = 28
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col_idx, value=h)
        cell.font = FONT_HEADER
        cell.fill = FILL_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_THIN
        
    # Data Rows
    start_row = 5
    for row_idx, r in enumerate(records, start=start_row):
        ws.row_dimensions[row_idx].height = 22
        is_even = (row_idx % 2 == 0)
        default_fill = FILL_ZEBRA if is_even else FILL_WHITE
        fill_to_use = FILL_STRUCK if r.get("struck") else default_fill
        font_to_use = FONT_STRUCK if r.get("struck") else FONT_DATA
        
        row_values = [
            r["sno"],
            r.get("block", ""),
            r["name"],
            r["role"],
            r["mobile"],
            r["location"],
            r["father"],
            r["remarks"],
            r.get("source", "")
        ]
        
        for col_idx, val in enumerate(row_values, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=str(val) if val != "" else "")
            cell.font = font_to_use
            cell.fill = fill_to_use
            cell.border = BORDER_THIN
            
            # Alignment
            if col_idx in (1, 5):  # S.No or Mobile
                cell.alignment = ALIGN_CENTER
            elif col_idx in (2, 8, 9):  # Block, Remarks, Source
                cell.alignment = ALIGN_CENTER
            else:
                cell.alignment = ALIGN_LEFT
                
    # Auto-fit column widths
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            # Skip title merged row
            if cell.row in (1, 2):
                continue
            val_str = str(cell.value or '')
            # approximate length for Indic scripts
            char_len = len(val_str.encode('utf-8')) / 2 if any(ord(c) > 127 for c in val_str) else len(val_str)
            if char_len > max_len:
                max_len = char_len
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

# 1. Create Master / Consolidated Sheet
ws_master = wb.create_sheet(title="Master - सभी ब्लॉक (All Blocks)")
master_records = []
master_sno = 1
for b in blocks_data:
    b_records = b.get("records")
    if isinstance(b_records, list):
        for rec in b_records:
            if isinstance(rec, dict):
                m_rec = dict(rec)
                m_rec["sno"] = master_sno
                m_rec["block"] = str(b.get("block_name", ""))
                m_rec["source"] = str(b.get("source", ""))
                master_records.append(m_rec)
                master_sno += 1

style_sheet(
    ws_master,
    "📊 मास्टर डेटा शीट - नालंदा एवं नवादा सभी प्रखंड (Consolidated Master List)",
    HEADERS,
    master_records,
    is_master=True
)

# 2. Create Individual Block Sheets
for b in blocks_data:
    raw_title = str(b.get("sheet_title", ""))
    safe_title = raw_title.replace("/", "-").replace(":", "")[:31]
    
    ws = wb.create_sheet(title=safe_title)
    
    block_records = []
    b_records = b.get("records")
    if isinstance(b_records, list):
        for rec in b_records:
            if isinstance(rec, dict):
                b_rec = dict(rec)
                b_rec["block"] = str(b.get("block_name", ""))
                b_rec["source"] = str(b.get("source", ""))
                block_records.append(b_rec)
        
    style_sheet(
        ws,
        f"📍 प्रखंड: {b['block_name']} (जिला: {b['district']})",
        HEADERS,
        block_records,
        is_master=False
    )

# Save workbook
output_path = r"c:\Users\alamn\Downloads\25 Sept\Digitalized_Block_Data.xlsx"
wb.save(output_path)
print(f"SUCCESS: Workbook saved with {len(wb.sheetnames)} sheets to {output_path}")
