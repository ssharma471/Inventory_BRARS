# Month-End Inventory App

This version uses the complete master inventory list and restricts inventory categories to exactly:

- Finished
- Packaging
- Supplies
- OIC

All existing products that previously used another category are placed under **Finished** so every product remains included and category totals always equal the grand total.

## Run the app

```bash
npm install
npm run dev
```

Open the local address shown in Terminal, usually `http://localhost:5173`.

## Open without npm installation

The production build is included in `dist`:

```bash
python3 -m http.server 8080 --directory dist
```

Then open `http://localhost:8080`.
