# Transit Proposal Defense — PowerPoint Generator

Generates a ready-to-present **PowerPoint (.pptx)** for the minor project proposal
defense:

> **An Intelligent Public Transit Route Recommendation System for Kathmandu Valley**
> IOE, Thapathali Campus — Department of Electronics and Computer Engineering

The deck follows the official **IOE Thapathali "Slide Design Guidelines"**
(Dinesh Baniya Kshatri, 2020) and the **12-minute Proposal Defense** structure.

---

## How to generate the slides

1. Make sure you have **Python 3.8+** installed.
2. Open a terminal in this folder and run:

   ```bash
   pip install -r requirements.txt
   python generate_slides.py
   ```

3. The file **`Proposal_Defense.pptx`** will be created in this folder.
   Double-click it to open in **Microsoft PowerPoint** and edit freely.

---

## Adding your own diagrams (optional but recommended)

Three slides use placeholders for images you said you'd edit yourself.
If you place these image files in the **same folder** before running the script,
they are embedded automatically:

| File name           | Used on slide                         |
|---------------------|---------------------------------------|
| `block_diagram.png` | Methodology: System Architecture      |
| `flowchart.png`     | Methodology: Algorithm & Flowchart    |
| `gantt.png`         | Tentative Timeline (Gantt Chart)      |

If a file is missing, the slide shows a clean labeled placeholder box instead,
so the deck is always complete. You can also just drag your images into those
slides directly inside PowerPoint after generating.

---

## Slide order (17 slides)

1. Cover
2. Presentation Outline
3. Motivation
4. Objectives
5. Scope of Project
6. Methodology: System Architecture (block diagram)
7. Methodology: Working Principle
8. Methodology: Algorithm & Flowchart
9. Methodology: Instrumentation (HW/SW)
10. Expected Results
11. Expected Results (Contd.)
12. Project Applications
13. Tentative Timeline (Gantt Chart)
14. Estimated Project Budget (table)
15. References
16. References (Contd.)
17. Thank You

---

## Styling enforced (per guideline)

- White background, black text, **no animations**
- **Cover:** Title Arial Bold 44 · Members Arial Bold 22 · Dept/Campus/Date Arial 20 · no footer
- **Titles:** Arial Bold, centered, 36–40 pt
- **Body:** Arial, left-justified, 28 pt main bullet / 24 pt sub-bullet
- ≤ 6 main bullets, ≤ 8 words per bullet, ≤ 2–3 sub-bullets
- **Footer** on every non-cover slide: date (left) + slide number (right)
- One figure / table per slide
