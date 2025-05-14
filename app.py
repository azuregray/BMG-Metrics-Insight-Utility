# app.py
import os, subprocess, json, time, sys, streamlit as st
from UnitValuesRendering import harvesterFunc, excelProcessor

def select_folder(key):
    if st.button(f"Set {key.replace('_', ' ').title()}", key=f"{key}_button"):
        result = subprocess.run(["python", "folder_selector.py"], capture_output=True, text=True)
        if result.returncode == 0:
            folder_data = json.loads(result.stdout)
            folder_path = folder_data.get("folder_path")
            if folder_path and os.path.isdir(folder_path):
                st.session_state[key] = folder_path
            else:
                st.session_state[key] = None
        else:
            st.error("Error selecting folder")

def open_file(file_path):
    try:
        if sys.platform == "win32":
            os.startfile(file_path)
        elif sys.platform == "darwin":
            subprocess.run(["open", file_path])
        else:
            subprocess.run(["xdg-open", file_path])
    except Exception as e:
        st.error(f"Could not open the file :: {e}")

def main():
    st.title("BMG Values Comparison Tool")

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

    select_folder("input_folder")
    select_folder("output_folder")

    st.write("**Input folder:**", st.session_state.input_folder if st.session_state.input_folder else "Not selected")
    st.write("**Output folder:**", st.session_state.output_folder if st.session_state.output_folder else "Not selected")

    if st.session_state.input_folder and st.session_state.output_folder:
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
        st.text_area("Error", st.session_state.error_message, height=200)

if __name__ == "__main__":
    main()