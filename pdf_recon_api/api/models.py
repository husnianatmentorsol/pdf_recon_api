from django.db import models

class ReconciliationRecord(models.Model):
    processed_at = models.DateTimeField(auto_now_add=True)
    min_date = models.DateField()
    max_date = models.DateField()
    total_transactions = models.IntegerField(default=0)
    bank_filename = models.CharField(max_length=255, blank=True, null=True)
    hotel_filename = models.CharField(max_length=255, blank=True, null=True)

    def __str__(self):
        return f"Recon {self.min_date} to {self.max_date} ({self.total_transactions} txns)"
