import streamlit as st
from utils import *

def main():
    st.set_page_config(page_title="Invoice Extraction Bot.........")
    st.title("Invoice Extraction Bot")
    st.subheader('I can extract invocie details from outlook')

    submit = st.button("Run")

    if submit:
        with st.spinner('Wait while i get the data for you....'):
            mail = MailSearch()
            df = mail.getExtractedData()

            if not df.empty:
                st.write(df.head())
                st.download_button(
                    "Download data as CSV", 
                    df.to_csv(index=False).encode("utf-8"),
                    "extracted_invoices.csv",
                    "text/csv",
                    key="download-csv",
                )
            else:
                st.info("No invoice data found in your emails.")
        
        st.success('Your data has been downloaded')


if __name__ == '__main__':
    main()