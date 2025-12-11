
# Secret Santa Optimized VBA Macro

This repository contains a **VBA macro for Excel** that generates Secret Santa pairings while optimizing for:

-  No self-assignments
-  Avoiding same-course pairings (soft constraint)
-  Maximizing complementary interests
-  Multiple iterations to find the best pairing

---

## How to Use

1. Open `SecretSantaOptimized.xlsm`.
2. Fill in the **Participants** sheet:

| First Name | Last Name | Email ID | Course | Interest 1 | Interest 2 | Interest 3 |
|------------|----------|----------|--------|------------|------------|------------|

3. Open VBA Editor (`Alt + F11`) → Insert → Module → Paste `SecretSantaOptimized.bas`.
4. Close VBA Editor.
5. Run the macro:
   - Press `Alt + F8`
   - Select `SecretSantaOptimized`
   - Click **Run**
6. Results appear in the **Results** sheet:

| Giver Name | Email ID | Receiver Name |

---

## Optional Features

- Can track **interest match score** for each pairing.
- Adjustable **iterations** to improve pairing quality.

---

## License

MIT License
