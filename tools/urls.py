# tools/urls.py
from django.urls import path
from . import views  # views/dashboard.py se import kiya jaaye

urlpatterns = [
    path('', views.dashboard, name='dashboard'),  # Home page
    # Future tools:
    path('pdf-text-extractor/', views.pdf_text_extractor, name='pdf_text_extractor'),
    path('pdf-merger/', views.pdf_merger, name='pdf_merger'),
    path('pdf-to-word/', views.pdf_to_word, name='pdf_to_word'),
    path('pdf-to-excel/', views.pdf_to_excel, name='pdf_to_excel'),
]
