# Smart Plagiarism Detection System

An AI-powered plagiarism detection platform that analyzes uploaded documents, detects semantic similarities, and generates intelligent plagiarism insights using NLP and Machine Learning techniques.

---

## Overview

The Smart Plagiarism Detection System is designed to identify content similarity beyond exact keyword matching. The project uses Natural Language Processing (NLP), semantic embeddings, and similarity analysis to compare documents and detect potential plagiarism.

The system allows users to:

* Upload documents for plagiarism analysis
* Extract and preprocess text from PDFs
* Compare semantic similarity between documents
* Generate plagiarism scores and similarity insights
* View results through a clean web interface

---

## Features

* AI-powered semantic plagiarism detection
* PDF document text extraction
* NLP-based text preprocessing
* Cosine similarity comparison
* Fast document analysis
* User-friendly frontend interface
* Flask backend integration
* Real-time plagiarism scoring

---

## Tech Stack

### Frontend

* HTML
* CSS
* JavaScript

### Backend

* Python
* Flask
* Flask-CORS

### AI / ML / NLP

* Sentence Transformers
* Scikit-learn
* Transformers
* PyTorch
* NumPy

### Database

* MongoDB

---

## Project Structure

```bash
PlagiarismDetection/
│
├── backend/
│   ├── app.py
│   └── ...
│
├── frontend/
│   ├── index.html
│   ├── home.html
│   └── ...
│
├── uploads/
├── requirements.txt
├── .gitignore
└── README.md
```

---

## Installation & Setup

### 1. Clone the Repository

```bash
git clone https://github.com/Gayathrii1276/PlagiarismDetection.git
```

### 2. Navigate to the Project Folder

```bash
cd PlagiarismDetection
```

### 3. Create Virtual Environment

```bash
python -m venv venv
```

### 4. Activate Virtual Environment

#### macOS/Linux

```bash
source venv/bin/activate
```

#### Windows

```bash
venv\Scripts\activate
```

### 5. Install Dependencies

```bash
pip install -r requirements.txt
```

---

## Running the Application

### Start Backend Server

```bash
cd backend
python app.py
```

The Flask server will start locally.

---

## How It Works

1. User uploads documents
2. System extracts text from PDFs
3. NLP preprocessing is applied
4. Sentence embeddings are generated
5. Cosine similarity is calculated
6. Plagiarism percentage and similarity insights are displayed

---

## AI & NLP Concepts Used

* Semantic Similarity
* Sentence Embeddings
* Cosine Similarity
* Natural Language Processing
* Document Vectorization
* Text Extraction

---

## Future Enhancements

* Multi-document comparison
* Advanced plagiarism visualization
* AI-generated plagiarism explanations
* User authentication system
* Cloud deployment
* Support for DOCX and TXT files
* Report generation and export

---

## Screenshots

Add your project screenshots here.

Example:

```bash
screenshots/homepage.png
screenshots/results.png
```

---

## Use Cases

* Academic plagiarism detection
* Assignment verification
* Content originality checking
* Research paper similarity analysis
* Educational institutions

---

## Author

### Gayathri Kommineni

Aspiring AI/ML Engineer passionate about building intelligent systems using Generative AI, NLP, and Machine Learning.

GitHub: [https://github.com/Gayathrii1276](https://github.com/Gayathrii1276)

---

## License

This project is for educational and learning purposes.

---

## Repository Link

[https://github.com/Gayathrii1276/PlagiarismDetection](https://github.com/Gayathrii1276/PlagiarismDetection)
