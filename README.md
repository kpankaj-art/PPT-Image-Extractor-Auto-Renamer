# 📸 PPT Image Extractor & Auto-Renamer

एक शानदार Streamlit app जो PowerPoint presentations से automatically images extract करके smart naming देता है।

## ✨ Features

- ✅ **Right-Side Image Extraction**: Slide के दाहिनी ओर (Far View) की सभी images automatically निकालता है
- ✅ **Smart Auto-Naming**: Slide के metadata के आधार पर images को automatic rename करता है
- ✅ **Metadata Watermarking**: Optional watermark के साथ outlet name और contact number add करता है
- ✅ **Batch Processing**: Multiple slides को एक साथ process करता है
- ✅ **Progress Tracking**: Real-time progress bar के साथ processing status दिखाता है
- ✅ **Error Handling**: Detailed error messages और graceful error recovery
- ✅ **Detailed Reports**: Extracted images का complete summary और statistics

## 🚀 Installation

### Requirements
- Python 3.8+
- pip

### Setup करें:

```bash
# Repository clone करें
git clone https://github.com/kpankaj-art/PPT-Image-Extractor-Auto-Renamer.git
cd PPT-Image-Extractor-Auto-Renamer

# Dependencies install करें
pip install -r requirements.txt
```

## 📖 Usage

### App को चलाएं:

```bash
streamlit run app.py
```

फिर अपने browser में खुलने वाली window में:

1. **PPTX File Upload करें** - अपनी PowerPoint presentation को upload करें
2. **Settings चुनें** - Sidebar से watermark option select करें (optional)
3. **"Extract & Rename करें" बटन दबाएं** - Processing शुरू होगी
4. **ZIP Download करें** - सभी extracted images को download करें

## 📋 Metadata Extraction

App automatically निम्नलिखित metadata को slide से निकालता है:

```
Outlet Name: [Company/Store Name]
Contact: [10-digit Mobile Number]
Type: [Media Type - Billboard, Standee, etc.]
Size: [Dimensions - e.g., 10x5]
```

### Example Format:

**Input Slide Text:**
```
Outlet Name: Flipkart
Contact: 9876543210
Type: Billboard
Size: 10x5
```

**Output Filename:**
```
FLIPKART_9876543210_BILLBOARD_10x5_1.png
```

## 🎨 Features Details

### 1. **Right-Side Image Filtering**
- सिर्फ slide के दाहिनी ओर (50% से आगे) की images extract होती हैं
- यह far-view/product images के लिए बेहतरीन है

### 2. **Automatic Watermarking**
- Image के top-left corner में outlet name और contact number add होता है
- Semi-transparent background के साथ readable text

### 3. **Batch Processing**
- Multiple slides को एक साथ process करता है
- Progress bar से real-time status दिखता है

### 4. **Smart Reporting**
- Extracted images का detailed table
- Outlet-wise और type-wise statistics
- Download करने के लिए timestamped ZIP file

## 📁 Output Format

```
outlet_images_20240909_143022.zip
├── FLIPKART_9876543210_BILLBOARD_10x5_1.png
├── FLIPKART_9876543210_BILLBOARD_10x5_2.png
├── AMAZON_9123456789_STANDEE_5x3_1.png
└── AMAZON_9123456789_STANDEE_5x3_2.png
```

## ⚙️ Configuration

### Settings (Sidebar में):

- **✅ Image पर Metadata Watermark add करें** - Watermark toggle करने के लिए

## 🐛 Troubleshooting

### "PPT file को read करने में error"
- File correctly formatted है क्या चेक करें
- File को reopen करके save करें (MS Word corruptions को remove करने के लिए)

### "Images नहीं मिली"
- Ensure करें कि slide में right-side पर images हैं
- Slide को edit करके images को check करें

### "Watermark दिखाई नहीं दे रहा"
- Settings में watermark option enable है क्या check करें
- Font files system पर install हैं क्या verify करें

## 📊 System Requirements

- **RAM**: Minimum 2GB
- **Disk**: 500MB free space
- **Internet**: Not required (fully offline works)
- **OS**: Windows, macOS, Linux

## 🔐 Data Privacy

- ✅ सभी processing locally होती है
- ✅ कोई data upload नहीं होता
- ✅ Completely offline tool
- ✅ Files automatically delete हो जाती हैं processing के बाद

## 📝 License

This project is open source and available under the MIT License.

## 👨‍💻 Author

- **Pankaj Kola** - kpankaj-art

## 💬 Support

किसी भी issue या suggestion के लिए GitHub issues में report करें।

## 🤝 Contributing

Contributions, issues, and feature requests are welcome! Feel free to check the issues page.

---

**Made with ❤️ for efficient media extraction and management**
