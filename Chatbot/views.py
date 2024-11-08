from django.shortcuts import render, redirect
from django.contrib.auth.forms import AuthenticationForm, UserCreationForm
from django.contrib.auth import authenticate, login, logout
from django.contrib import messages
from .chatbot_model_python import generate_answer,add_docs,remove_docs,get_relevant_documents

import os
import pandas as pd
from django.conf import settings

from django.http import StreamingHttpResponse

from django.http import JsonResponse


def accueil_chatbot(request):
    return render(request,"Chatbot/chatbot_accueil.html",{})



def documents_chatbot(request):
    # Path to the directory containing the documents
    doc_folder = 'C:/Users/guill/OneDrive/Documents/HOPE/Test_django/gestion_ressources_V2/Chatbot/Doc traités'
    
    # List the files in the directory
    document_list = []
    if os.path.exists(doc_folder):
        document_list = os.listdir(doc_folder)

    # Pass the file list to the template
    return render(request, "Chatbot/gestion_doc_chatbot.html", {"document_list": document_list})

def documents_chatbot2(request):
    # Path to the directory containing the documents
    doc_folder = 'C:/Users/guill/OneDrive/Documents/HOPE/Test_django/gestion_ressources_V2/Chatbot/Doc traités'
    
    # List the files in the directory
    document_list = []
    if os.path.exists(doc_folder):
        document_list = os.listdir(doc_folder)

    # Pass the file list to the template
    return render(request, "Chatbot/gestion_doc_chatbot_test.html", {"document_list": document_list})




def ajout_document_template(request):
    return render(request, "Chatbot/ajout_doc.html", {})



from django.core.files.storage import default_storage


def ajout_document_fonction(request):
    if request.method == 'POST':
        uploaded_file = request.FILES['document']  # Get the uploaded file
        file_path = os.path.join(settings.MEDIA_ROOT, uploaded_file.name)

        # Save the file to the media directory
        with default_storage.open(file_path, 'wb+') as destination:
            for chunk in uploaded_file.chunks():
                destination.write(chunk)

        # Call the add_docs function with the saved file path
        add_docs(file_path, treated_folder='C:/Users/guill/OneDrive/Documents/HOPE/Test_django/gestion_ressources_V2/Chatbot/Doc traités')

        return redirect('Chatbot:documents_chatbot')  # Redirect after successful upload

    return render(request, 'ajout_doc.html')
    
def remove_doc(request, document_name):
    doc_folder = 'C:/Users/guill/OneDrive/Documents/HOPE/Test_django/gestion_ressources_V2/Chatbot/Doc traités'
    doc_path = os.path.join(doc_folder, document_name)

    if os.path.exists(doc_path):
        remove_docs([doc_path])  # Call the function to remove the document
    return redirect('Chatbot:documents_chatbot')


import re
import time
from django.http import StreamingHttpResponse
from .chatbot_model_python import generate_answer

def format_response_text(text):
    """
    Format the GPT-4 response to handle newlines, tabs, bold, code blocks, etc.
    """
    # Replace newlines with <br> for HTML
    text = text.replace('\n', '<br>')

    # Replace tabs with four non-breaking spaces
    text = text.replace('\t', '&nbsp;&nbsp;&nbsp;&nbsp;')

    # Handle bold text (Markdown style: **bold**)
    text = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', text)

    # Handle code blocks (```code```)
    text = re.sub(r'```(.*?)```', r'<pre><code>\1</code></pre>', text, flags=re.DOTALL)

    # Handle inline code (Markdown style: `code`)
    text = re.sub(r'`(.*?)`', r'<code>\1</code>', text)

    return text

def chatbot_application(request):
    if request.method == 'POST':
        user_input = request.POST.get('user_input')

        # Call the function that generates the response from GPT-4
        answer = generate_answer(user_input)

        # Format the GPT-4 response before streaming
        formatted_answer = format_response_text(answer)

        def stream_response():
            """Stream the response word by word"""
            words = formatted_answer.split()
            for word in words:
                yield word + " "  # Stream each word with formatting
                time.sleep(0.05)  # Simulate typing delay

        return StreamingHttpResponse(stream_response(), content_type='text/html')

    return render(request, "Chatbot/chatbot_test.html", {})


from django.http import JsonResponse

def get_sources(request):
    if request.method == 'POST':
        query = request.POST.get('query', '')
        relevant_docs = get_relevant_documents(query)  # Assuming you already have this function
        top_sources = [{'segment': f"<p>{doc['segment']}</p>", 'metadata': doc['metadata']} for doc in relevant_docs[:3]]  # Get the top 3 sources and format them as HTML paragraphs
        return JsonResponse({'sources': top_sources})
    
    return JsonResponse({'error': 'Invalid request method'}, status=400)







