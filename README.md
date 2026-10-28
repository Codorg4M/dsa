# 🎬 Movie Rating Sorter Using Bubble Sort
**Data Structures and Algorithms (DSA) Laboratory Mini Project**  
*Bachelor of Technology (B.Tech) in Computer Science & Engineering*

---

## 📌 Project Overview
The **Movie Rating Sorter** is an interactive, educational web application designed to demonstrate the **Bubble Sort** algorithm applied to real-world data (movie ratings).

The project sorts and ranks movies (e.g., *The Dark Knight*, *Inception*, *Interstellar*, *3 Idiots*, *Dangal*) based on their IMDb-style ratings (0.0 to 10.0), featuring real-time visual step-by-step animations of adjacent comparisons, swaps, and settled positions.

---

## 📁 Project File Structure
```text
Dsa Project/
│
├── index.html                           # Complete semantic UI (Header, Theory, Visualizer, Viva)
├── style.css                            # Modern Cinema Dark Theme, Grid layout, and animations
├── script.js                            # Core Bubble Sort algorithm, visualizer, & event logic
├── Movie_Rating_Sorter_Project_Report.docx  # Full Academic B.Tech Project Report (Word Document)
├── generate_report.py                   # Python script used to generate the Word report
└── README.md                            # Documentation and viva presentation guide
```

---

## 🫧 Key DSA Concepts Implemented
1. **Manual Bubble Sort Implementation:**  
   Strictly zero reliance on built-in methods like `Array.prototype.sort()`.
2. **Bidirectional Sorting:**
   - **Highest → Lowest (Descending):** Swaps when `rating[j] < rating[j + 1]` (IMDb top-rated view).
   - **Lowest → Highest (Ascending):** Swaps when `rating[j] > rating[j + 1]`.
3. **Early Exit Optimization:**  
   Uses a boolean `swapped` flag. If no swaps occur in a pass, the algorithm terminates early, reducing best-case time complexity to **$O(n)$**.
4. **Time & Space Complexity:**
   - **Best Case:** $O(n)$ (already sorted with optimization flag)
   - **Average Case:** $O(n^2)$ (arbitrary random order)
   - **Worst Case:** $O(n^2)$ (reverse sorted order)
   - **Space Complexity:** $O(1)$ (In-place algorithm)

---

## 🚀 How to Run the Project
1. **Directly in Browser:**  
   Double-click `index.html` or right-click `index.html` and select **Open with Google Chrome / Microsoft Edge / Firefox**.
2. **No Setup Required:**  
   Works 100% offline with zero dependencies or server installations.

---

## 📄 Word Document Report Details
The generated file `Movie_Rating_Sorter_Project_Report.docx` contains a formal academic report formatted for college submission:
- **Title Page / Cover Page** with University, Department, and Student details
- **Certificate of Authenticity** with Guide, HOD, and External Examiner signature blocks
- **Candidate Declaration & Acknowledgement**
- **Abstract** with key project highlights
- **Table of Contents** with 7 full chapters
- **Chapter 1: Introduction & Problem Statement**
- **Chapter 2: Theoretical Background & Algorithm Details** (with 7-step pseudocode, dry run, and complexity matrices)
- **Chapter 3: System Design & Architecture** (SRS, Flowchart, Data representation)
- **Chapter 4: Implementation Details & Code Walkthrough** (script.js explanation)
- **Chapter 5: Testing, Experimental Results & Analysis** (Test cases & performance logs)
- **Chapter 6: College Viva Voce Top 10 Q&A**
- **Chapter 7: Conclusion & Future Scope**
- **References & Bibliography**

---

## 🎓 How to Present During a College Viva
1. **Open `index.html`** in your browser.
2. Point out the **Theory Section** explaining adjacent comparisons, time complexities ($O(n)$ best, $O(n^2)$ worst), and $O(1)$ space.
3. Show the **Sample Movie List** (8 titles preloaded).
4. Click **"⬇️ Sort Highest → Lowest"**:
   - Highlight the **Cyan border** (`.comparing`) showing adjacent elements being compared.
   - Point out the **Red badge** (`.swapping`) when elements are swapped.
   - Point out the **Emerald green badge** (`.sorted`) as elements settle into their final sorted positions at the end of each pass.
5. Demonstrate the **Early Exit Optimization**:
   - Click sort again on the already-sorted list.
   - Show how the visualizer finishes after **Pass 1** in $O(n)$ time with 0 swaps!
6. Refer to the **Viva Voce section** at the bottom of the page or Chapter 6 in the Word report for examiner questions.
