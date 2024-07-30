import sys
import os
import pathlib
import shutil
import subprocess


def compile_tailwind():
    try:
        # Assuming that your Tailwind CSS file is `./src/tailwind.css`
        # and you want to output to `./build/tailwind.css`.
        command = "npx tailwindcss -i css/style.css -o build/style.css"
        process = subprocess.Popen(command, shell=True, stdout=subprocess.PIPE)
        process.wait()
        print("Command executed successfully. Exit code:", process.returncode)
    except subprocess.CalledProcessError as e:
        print("An error occurred while executing the command. Error: ", e)


def main():

    # delete dir "build"
    shutil.rmtree("build", ignore_errors=True)

    # make dir "build"
    pathlib.Path("build").mkdir(parents=True, exist_ok=True)

    # copy files
    shutil.copy("src/index.html", "build/index.html")
    shutil.copy("src/script.js", "build/script.js")

    # compile tailwind css
    compile_tailwind()

    # copy other files
    shutil.copytree("src/logos", "build/logos")
    shutil.copytree("src/assets", "build/assets")
    shutil.copytree("src/download", "build/download")

    # shutil.copy("src/products/seznam-náhradního-spotřebního-materiálu.xlsx", "build/download/seznam-náhradního-spotřebního-materiálu.xlsx")
    # shutil.copy("src/products/seznam-náhradního-spotřebního-materiálu.txt",  "build/download/seznam-náhradního-spotřebního-materiálu.txt" )
    
    shutil.copy("src/products/seznam-náhradního-spotřebního-materiálu.zip", "build/download/seznam-náhradního-spotřebního-materiálu.zip")

    shutil.copytree("src/products/product-images", "build/product-images")

    shutil.copy("CNAME", "build/CNAME")



if __name__ == "__main__":
    main()
