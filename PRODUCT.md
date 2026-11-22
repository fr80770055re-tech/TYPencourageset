# Product

<!-- impeccable:product-schema 1 -->

## Platform

web

## Users
- **Teachers** (homeroom, subject, admin; `typt…` school Google accounts). They give bonus points between lessons, on a phone and on a classroom computer about equally often.
- **Students** (grade 5 pilot, `typs<學號>` accounts). They check their own total and history, mostly on school tablets or Chromebooks.

## Product Purpose
高雄市桃源國小全校加分系統 replaces paper bonus slips that were cut and glued into contact books. Success means a teacher can give points in a few taps, and nobody has to cut or glue anything.

## Operating Context
- Pilot class: 五年忠班, 14 students. The school-wide rollout is approved as the long-term goal.
- One URL. The role is decided by the school account; nobody logs in separately.
- The classroom computer is **sometimes projected** to the whole class. The class view must not expose every student's points by default.
- Homeroom teachers copy monthly totals into a school reporting form; monthly totals are settled on the 1st.

## Capabilities and Constraints
- Google Apps Script web app (HtmlService, sandboxed iframe). The front end is a single `Index.html`, with no build step.
- Native `confirm()`/`alert()` dialogs can be blocked inside the iframe. Use in-page confirmation instead.
- Points come in steps of 10, with no upper limit. The common reasons are configurable.
- Withdrawn records are marked, never deleted.

## Brand Commitments
- Name: **高雄市桃源國小全校加分系統**.
- Palette pinned by the user: low-chroma red, green and yellow, with a frosted-glass material and animated transitions on every tap.
- The mobile and desktop layouts must differ, not just stretch.

## Product Principles
1. Most teachers must accept it: fewest steps, no training needed.
2. Never embarrass a student in front of the class.
3. Records are trustworthy: visible who gave what, and withdrawals leave a trace.
