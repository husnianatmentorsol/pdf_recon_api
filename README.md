# Reconcile API

Backend service for reconciling bank and hotel settlement files and generating
reports (Excel + HTML preview).

## Quick start

1) Create a virtual environment and install dependencies (see `requirements.txt` if present).
2) Run the Django server from the project root.
3) POST to `/api/reconcile/` with multipart form data:

```
bank_file: [file]
hotel_file: [file]
client_name: "Client A"
threshold_time: 15
```

## Notes

- Uploaded files and generated reports are stored under `media/`.
- Google credential JSONs are ignored by git; configure them locally as needed.
