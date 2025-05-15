import os, streamlit as st

def uploadsHandler(session_key):
    uploaded_InputFiles = st.file_uploader("Upload all files from the Input folder...", accept_multiple_files=True, key="input_uploader")
    uploaded_OutputFiles = st.file_uploader("Upload all files from the Output folder...", accept_multiple_files=True, key="output_uploader")

    INPUT_FOLDER = f'InputFilesFolder_{session_key}'
    OUTPUT_FOLDER = f'OutputFilesFolder_{session_key}'
    
    if uploaded_InputFiles and uploaded_OutputFiles:
        os.makedirs(INPUT_FOLDER, exist_ok=True)  # Create the folder if it doesn't exist
        for file in uploaded_InputFiles:
            file_path = os.path.join(INPUT_FOLDER, file.name)
            with open(file_path, "wb") as f:
                f.write(file.read())
        
        os.makedirs(OUTPUT_FOLDER, exist_ok=True)  # Create the folder if it doesn't exist
        for file in uploaded_OutputFiles:
            file_path = os.path.join(OUTPUT_FOLDER, file.name)
            with open(file_path, "wb") as f:
                f.write(file.read())
        return INPUT_FOLDER, OUTPUT_FOLDER
    return None, None

def cleaningHandler(session_key):
    import shutil
    
    for item in os.listdir():
        if session_key in item:
            try:
                shutil.rmtree(item)
            except:
                print(f'There was a problem while clearning folders for session_key: {session_key}')

def main():
    import secrets, time
    from UnitValuesRendering import harvesterFunc, excelProcessor
    
    st.title("BMG Values Comparison Tool")

    if "session_key" not in st.session_state:
        st.session_state.session_key = secrets.token_hex(4)
    if "input_folder" not in st.session_state:
        st.session_state.input_folder = None
    if "output_folder" not in st.session_state:
        st.session_state.output_folder = None
    if "comparison_results" not in st.session_state:
        st.session_state.comparison_results = None
    if "backend_results" not in st.session_state:
        st.session_state.backend_results = None
    if "error_message" not in st.session_state:
        st.session_state.error_message = None
    if "excel_output_path" not in st.session_state:
        st.session_state.excel_output_path = None
    if "time_taken" not in st.session_state:
        st.session_state.time_taken = None
    
    session_key = st.session_state.session_key
    
    if not st.session_state.input_folder or not st.session_state.output_folder:
        st.session_state.input_folder, st.session_state.output_folder = uploadsHandler(session_key)

    if st.session_state.input_folder and st.session_state.output_folder:
        st.write("**✅ Your files have been saved successfully.**")
        if st.button("Run Comparison"):
            st.subheader("Processing...")

            progress = st.progress(0)
            percent_text = st.empty()

            def update_progress(percent):
                progress.progress(percent)
                percent_text.markdown(f"**Progress: {int(percent * 100)}%**")

            try:
                startTime = time.perf_counter()
                
                st.session_state.backend_results = harvesterFunc(
                    st.session_state.input_folder,
                    st.session_state.output_folder,
                    progress_callback=update_progress
                )

                endTime = time.perf_counter()
                st.session_state.time_taken = endTime - startTime
                
                cleaningHandler(session_key)

                output_lines = []
                for filename, values_in, values_out, match in st.session_state.backend_results:
                    if match == 'True':
                        status = "✅✅ Values Match ✅✅"
                        line = f"{filename}  ==>  {status}\nInput File Values = {values_in}\nOutput File Values = {values_out}"
                    elif match == 'False':
                        status = "❌❌ Values Do Not Match ❌❌"
                        line = f"{filename}  ==>  {status}\nInput File Values = {values_in}\nOutput File Values = {values_out}"
                    elif match == 'unmatched':
                        status = "⚠️⚠️ File Match Error ⚠️⚠️"
                        line = f"{filename}  ==>  {status}"
                    output_lines.append(line)
                st.session_state.comparison_results = "\n\n".join(output_lines)
                st.session_state.error_message = None
                update_progress(1.0) # Ensure progress is 100%

            except Exception as e:
                st.session_state.error_message = f"There was an error while processing your task :: {e}"
                st.session_state.comparison_results = None
                st.session_state.excel_output_path = None
                st.session_state.time_taken = None

    if st.session_state.comparison_results:
        st.text_area("Comparison Results", st.session_state.comparison_results, height=200)
        if st.session_state.time_taken is not None:
            st.markdown(f"**System took {st.session_state.time_taken:.2f} seconds to process your files.**")
        
        exportedPath = excelProcessor(
            exportTemplatePath = os.path.join(os.path.dirname(__file__), 'Results-ExportTemplate.xlsx'),
            exportableData = st.session_state.backend_results
        )
        
        with open(exportedPath, "rb") as exp:
            readCur = exp.read()
            st.download_button(label="Export Results to Excel",
                                data=readCur,
                                file_name=os.path.basename(exportedPath),
                                mime='application/octet-stream')

    if st.session_state.error_message:
        cleaningHandler(session_key)
        st.text_area("Error", st.session_state.error_message, height=200)

if __name__ == "__main__":
    main()