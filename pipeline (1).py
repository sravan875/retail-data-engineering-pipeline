
################################################ Paste your entire pipeline code here
from IPython.display import display
import ipywidgets as widgets

upload = widgets.FileUpload(
    accept=".xlsx",
    multiple=False
)

display(upload)

print("Upload USECASE_-_Data_Engineering.xlsx")
print("After upload run Cell 2")

################################################

uploaded_filename = list(upload.value.keys())[0]

with open(uploaded_filename, "wb") as f:

    f.write(
        upload.value[
            uploaded_filename
        ]["content"]
    )

DATA_FILE = uploaded_filename

OUTPUT_DIR="output"

print("Uploaded Successfully")

print(DATA_FILE)

################################################
import pandas as pd
import numpy as np
import hashlib
import re
import os
import logging

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s"
)

CATEGORY_MAP = {

"electronics":"Electronics",
"elec":"Electronics",

"clothing":"Clothing",
"cloth":"Clothing",

"furniture":"Furniture",
"furn":"Furniture",

"home appliances":"Home Appliances",
"home":"Home Appliances"

}

PRODUCT_MAP={

"phone":"Phone",
"tv":"TV",
"shirt":"Shirt",
"shoes":"Shoes",

"laptop":"Laptop",

"refrigerator":"Refrigerator",

"microwave":"Microwave",

"sofa":"Sofa"

}

################################################

def ingest_data(filepath):

    product_raw=pd.read_excel(

        filepath,

        sheet_name="PRODUCT DETAILS",

        header=None

    )

    product=product_raw.iloc[2:]

    product=product.reset_index(drop=True)

    product.columns=product.iloc[0]

    product=product[1:]

    product.columns=[

        "product_id",

        "product_name",

        "category",

        "price"

    ]

    r1_raw=pd.read_excel(

        filepath,

        sheet_name="RETAIL DATA 1",

        header=None

    )

    r1=r1_raw.iloc[1:]

    r1=r1.reset_index(drop=True)

    r1.columns=r1.iloc[0]

    r1=r1[1:]

    r2_raw=pd.read_excel(

        filepath,

        sheet_name="RETAIL DATA 2",

        header=None

    )

    idx=r2_raw[

        r2_raw.iloc[:,0]

        =="transaction_id"

    ].index[0]

    r2=r2_raw.iloc[idx:]

    r2=r2.reset_index(drop=True)

    r2.columns=r2.iloc[0]

    r2=r2[1:]

    columns=[

        "transaction_id",

        "customer_id",

        "customer_name",

        "product_id",

        "price",

        "product_name",

        "category",

        "purchase_location",

        "city",

        "transaction_date",

        "quantity",

        "payment_method",

        "discount",

        "email",

        "phone",

        "payment_status"

    ]

    r1.columns=columns

    r2.columns=columns

    return product,r1,r2

################################################

def clean_data(df,product_dim):

    nums=[

        "transaction_id",

        "customer_id",

        "product_id",

        "price",

        "quantity",

        "discount"

    ]

    for c in nums:

        df[c]=pd.to_numeric(

            df[c],

            errors="coerce"

        )

    df=df.dropna(

        subset=[

            "transaction_id",

            "customer_id"

        ]

    )

    df=df.drop_duplicates()

    df=df.drop_duplicates(

        subset=["transaction_id"]

    )

    df["category"]=df["category"]\
        .astype(str)\
        .str.lower()

    df["category"]=df["category"]\
        .map(CATEGORY_MAP)\
        .fillna(df["category"])

    df["product_name"]=df[
        "product_name"
    ].astype(str)\
    .str.lower()

    df["product_name"]=df[
        "product_name"
    ].map(PRODUCT_MAP)\
    .fillna(
        df["product_name"]
    )

    df["transaction_date"]=pd.to_datetime(

        df["transaction_date"],

        errors="coerce"

    )

    df=df.dropna(

        subset=["transaction_date"]

    )

    lookup=product_dim.set_index(

        "product_id"

    )["price"].to_dict()

    missing=df["price"].isna()

    df.loc[
        missing,
        "price"
    ]=df.loc[
        missing,
        "product_id"
    ].map(
        lookup
    )

    df["discount"]=df[
        "discount"
    ].fillna(0)

    df=df[
        df["quantity"]>0
    ]

    df=df[
        df["payment_status"]
        .str.lower()
        =="successful"
    ]

    df["gross_revenue"]=(
        df["price"]
        *
        df["quantity"]
    )

    df["discount_amount"]=(
        df["gross_revenue"]
        *
        df["discount"]
    )

    df["net_revenue"]=(
        df["gross_revenue"]
        -
        df["discount_amount"]
    )

    return df

################################################

def mask_email(x):

    if pd.isna(x):

        return "N/A"

    return hashlib.sha256(

        str(x).encode()

    ).hexdigest()[:16]

def mask_phone(x):

    digits=re.sub(

        r"\D",

        "",

        str(x)

    )

    if len(digits)<4:

        return "****"

    return "*"*(len(digits)-4)+digits[-4:]

################################################

def apply_masking(df):

    df["email_masked"]=df[
        "email"
    ].apply(mask_email)

    df["phone_masked"]=df[
        "phone"
    ].apply(mask_phone)

    df=df.drop(

        columns=[

            "email",

            "phone"

        ]

    )

    return df

################################################

def compute_kpis(df):

    summary=pd.DataFrame({

        "Metric":[

            "Net Revenue",

            "Transactions",

            "Units Sold"

        ],

        "Value":[

            df["net_revenue"].sum(),

            len(df),

            df["quantity"].sum()

        ]

    })

    return summary

################################################

def export(df,summary):

    os.makedirs(

        OUTPUT_DIR,

        exist_ok=True

    )

    df.to_csv(

        OUTPUT_DIR+"/retail_fact.csv",

        index=False

    )

    with pd.ExcelWriter(

        OUTPUT_DIR+"/retail_kpis.xlsx"

    ) as writer:

        df.to_excel(

            writer,

            sheet_name="fact",

            index=False

        )

        summary.to_excel(

            writer,

            sheet_name="summary",

            index=False

        )

################################################

def run_pipeline():

    product,r1,r2=ingest_data(

        DATA_FILE

    )

    raw=pd.concat(

        [r1,r2],

        ignore_index=True

    )

    clean=clean_data(

        raw,

        product

    )

    clean=apply_masking(

        clean

    )

    summary=compute_kpis(

        clean

    )

    export(

        clean,

        summary

    )

    print()

    print("PIPELINE COMPLETED")

    print()

    print(summary)

################################################

run_pipeline()