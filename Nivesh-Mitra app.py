# main.py

from fastapi import FastAPI, HTTPException
from pydantic import BaseModel, Field
from datetime import date
from typing import Optional, List

app = FastAPI(
    title="E-Way Bill Closure API",
    description="API middleware for GST E-Way Bill Ship-To validation and EWB closure",
    version="1.0.0"
)


class ShipToValidationRequest(BaseModel):
    transaction_type: str = Field(..., example="BILL_TO_SHIP_TO")
    ship_to_gstin: Optional[str] = Field(None, example="07ABCDE1234F1Z5")
    consignee_registered: bool = Field(..., example=True)


class EWayBillClosureRequest(BaseModel):
    ewb_no: str = Field(..., example="181001234567")
    closure_date: date
    remarks: str = Field(..., example="Goods delivered successfully")
    closed_by: str = Field(..., example="Supplier")


class BulkEWayBillClosureRequest(BaseModel):
    items: List[EWayBillClosureRequest]


@app.post("/api/ewaybill/validate-ship-to")
def validate_ship_to(data: ShipToValidationRequest):
    if data.transaction_type.upper() == "BILL_TO_SHIP_TO":
        if data.consignee_registered and not data.ship_to_gstin:
            raise HTTPException(
                status_code=400,
                detail="Ship To GSTIN is mandatory for registered consignee."
            )

        if not data.consignee_registered:
            if data.ship_to_gstin != "URP":
                raise HTTPException(
                    status_code=400,
                    detail='For unregistered consignee, Ship To GSTIN must be "URP".'
                )

    return {
        "status": "success",
        "message": "Ship To GSTIN validation passed"
    }


@app.post("/api/ewaybill/close")
def close_eway_bill(data: EWayBillClosureRequest):
    # Here we will call NIC/GSP API after official credentials are available.
    nic_payload = {
        "ewbNo": data.ewb_no,
        "closureDate": str(data.closure_date),
        "remarks": data.remarks
    }

    return {
        "status": "success",
        "message": "E-Way Bill closure request prepared successfully",
        "nic_payload": nic_payload,
        "note": "Replace mock response with actual NIC Sandbox/Production API call."
    }


@app.post("/api/ewaybill/bulk-close")
def bulk_close_eway_bill(data: BulkEWayBillClosureRequest):
    results = []

    for item in data.items:
        results.append({
            "ewb_no": item.ewb_no,
            "closure_date": str(item.closure_date),
            "remarks": item.remarks,
            "status": "prepared"
        })

    return {
        "status": "success",
        "total_records": len(results),
        "results": results
    }


@app.get("/api/ewaybill/closure-status/{ewb_no}")
def closure_status(ewb_no: str):
    return {
        "ewb_no": ewb_no,
        "closure_status": "Mock status - NIC API integration pending"
    }


@app.get("/api/audit-log")
def audit_log():
    return {
        "message": "Audit log module can be connected with SQLite/PostgreSQL"
    }
