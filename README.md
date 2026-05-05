# 🧵 Textile Inspection Hub — Node.js + MongoDB Atlas

Full-stack inspection report system with image upload and Excel export.

---

## 📁 Project Structure

```
textile-inspection/
├── config/
│   └── db.js                   # MongoDB Atlas connection
├── controllers/
│   └── inspectionController.js # All business logic + Excel export
├── middleware/
│   └── upload.js               # Multer image upload config
├── models/
│   └── Inspection.js           # Mongoose schema
├── routes/
│   └── inspections.js          # API routes
├── public/
│   ├── index.html              # Frontend (full form)
│   └── uploads/                # Uploaded images stored here
├── .env                        # Your secret config (edit this!)
├── .env.example                # Template
├── server.js                   # Entry point
└── package.json
```

---

## ⚙️ Setup Instructions

### Step 1 — Install Node.js
Download from https://nodejs.org (v18 or higher recommended)

### Step 2 — MongoDB Atlas Setup

1. Go to https://mongodb.com/atlas and create a **free account**
2. Create a **new cluster** (free M0 tier)
3. In "Database Access" → Add a new user (username + password)
4. In "Network Access" → Add IP: `0.0.0.0/0` (allow all, for development)
5. Click "Connect" → "Connect your application"
6. Copy the connection string — it looks like:
   ```
   mongodb+srv://myuser:mypassword@cluster0.abcde.mongodb.net/?retryWrites=true&w=majority
   ```

### Step 3 — Configure .env

Edit the `.env` file in the project root:

```env
PORT=3000
MONGODB_URI=mongodb+srv://YOUR_USER:YOUR_PASSWORD@YOUR_CLUSTER.mongodb.net/textile_inspection?retryWrites=true&w=majority
NODE_ENV=development
```

Replace `YOUR_USER`, `YOUR_PASSWORD`, and `YOUR_CLUSTER` with your actual values.

### Step 4 — Install Dependencies

Open terminal in the project folder and run:

```bash
npm install
```

### Step 5 — Start the Server

```bash
npm start
```

Or for development with auto-restart:
```bash
npm run dev
```

### Step 6 — Open in Browser

```
http://localhost:3000
```

---

## 🚀 Features

| Feature | Description |
|---|---|
| **Save to MongoDB** | Saves complete inspection form data to MongoDB Atlas |
| **Load Saved** | Browse, load, or delete previously saved inspections |
| **Image Upload** | Click any image card → upload image → saved to server |
| **Excel Export** | Downloads `.xlsx` with all form data + embedded images |
| **New Inspection** | Start fresh with default data |

---

## 🔌 API Endpoints

| Method | URL | Description |
|---|---|---|
| GET | `/api/inspections` | List all inspections |
| POST | `/api/inspections` | Create new inspection |
| GET | `/api/inspections/:id` | Get single inspection |
| PUT | `/api/inspections/:id` | Update inspection |
| DELETE | `/api/inspections/:id` | Delete inspection |
| POST | `/api/inspections/upload/image` | Upload a single image |
| GET | `/api/inspections/:id/excel` | Download Excel file |

---

## 📝 How to Use

1. **Fill in the form** — buyer, PO, factory, defects, etc.
2. **Upload images** — click any image card → select photo
3. **Click "Save to MongoDB"** — data saved to Atlas
4. **Click "Download Excel"** — gets latest data + images embedded in `.xlsx`
5. **Click "Load Saved"** — open any previous inspection

---

## 🛠️ Troubleshooting

**MongoDB connection fails:**
- Check your `.env` MONGODB_URI is correct
- Make sure your IP is whitelisted in Atlas Network Access
- Verify username/password in Atlas Database Access

**Images not showing in Excel:**
- Make sure you saved the inspection first before downloading Excel
- Image files must still exist in `public/uploads/`

**Port already in use:**
- Change `PORT=3000` in `.env` to another port like `3001`
