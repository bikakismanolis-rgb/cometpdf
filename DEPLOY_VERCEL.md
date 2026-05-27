# CometPDF: GitHub + Vercel deploy

Αυτή είναι η πιο απλή διαδρομή για να ανέβει το site online.

## 1. Φτιάχνεις repo στο GitHub

1. Πήγαινε στο GitHub και πάτα **New repository**.
2. Όνομα: `cometpdf` ή `cometpdf-site`.
3. Μπορεί να είναι **Private** στην αρχή.
4. Μην προσθέσεις README από το GitHub αν θα ανεβάσεις όλο αυτόν τον φάκελο.

## 2. Ανεβάζεις τα αρχεία

Αν χρησιμοποιείς GitHub Desktop:

1. **File > Add local repository**.
2. Διάλεξε τον φάκελο του project:
   `C:\Users\user\Documents\Codex\2026-05-25\pdf-pdf-pdf`
3. Αν σου πει ότι δεν είναι Git repo, πάτα **create a repository**.
4. Κάνε commit με μήνυμα `Initial CometPDF website`.
5. Πάτα **Publish repository**.

Το `.gitignore` κρατάει έξω τα προσωρινά build folders. Τα αρχεία του site και τα downloads μέσα στο `website/downloads` ανεβαίνουν κανονικά.

## 3. Το συνδέεις με Vercel

1. Πήγαινε στο Vercel.
2. **Add New > Project**.
3. Κάνε import το GitHub repo.
4. Στο **Root Directory** βάλε:
   `website`
5. Framework preset: **Other** ή **Static**.
6. Build command: άδειο.
7. Output directory: άδειο ή `.` αν στο ζητήσει.
8. Πάτα **Deploy**.

Μετά από λίγα δευτερόλεπτα θα σου δώσει ένα link τύπου:
`https://cometpdf-xxxxx.vercel.app`

## 4. Μετά το πρώτο deploy

Όταν έχεις κανονικό domain, π.χ. `cometpdf.com`, πρέπει να αλλάξουμε τα placeholder URLs:

- `website/sitemap.xml`
- canonical links μέσα στις σελίδες
- `website/robots.txt`

Μετά το ανεβάζεις ξανά και το δηλώνεις στο Google Search Console.

## Σημαντικό

Τα αρχεία εγκατάστασης είναι μεγάλα. Αν το Vercel δεν δεχτεί τα downloads ή αργότερα θέλουμε πιο επαγγελματική διανομή, η σωστή λύση είναι να βάλουμε τα `.exe` και `.zip` σε GitHub Releases και τα κουμπιά της σελίδας να δείχνουν εκεί.
