import io
import requests  # <-- Import library requests untuk komunikasi dengan server Google reCAPTCHA
from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.decorators import login_required
from decimal import Decimal
from django.utils import timezone
from django.http import HttpResponse
from openpyxl import Workbook
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from .models import Pasien, Kriteria
from .utils import hitung_cf

def login_view(request):
    # Jika user sudah login, langsung alihkan ke dashboard
    if request.user.is_authenticated:
        return redirect('dashboard')

    pesan_error = None
    
    if request.method == 'POST':
        username_input = request.POST.get('username', '')
        password_input = request.POST.get('password', '')
        
        # Tangkap token respon dari widget Google reCAPTCHA di HTML
        recaptcha_response = request.POST.get('g-recaptcha-response')
        
        secret_key = '6LcAWLwtAAAAAMkwEymPq6sMYABVvKppijHjahir'
        
        # Kirim data verifikasi token ke server Google API
        payload = {
            'secret': secret_key,
            'response': recaptcha_response
        }
        response = requests.post('https://www.google.com/recaptcha/api/siteverify', data=payload)
        result = response.json()
        
        # Cek apakah verifikasi sukses dari Google
        if not result.get('success'):
            pesan_error = "Validasi reCAPTCHA gagal! Silakan centang kotak verifikasi robot."
        else:
            # Jika reCAPTCHA valid, baru proses autentikasi username & password
            user = authenticate(request, username=username_input, password=password_input)
            if user is not None:
                login(request, user)
                return redirect('dashboard')
            else:
                pesan_error = "Username atau password salah! Silakan coba lagi."
            
    return render(request, 'login.html', {'pesan_error': pesan_error})

def logout_view(request):
    logout(request)
    return redirect('login')

@login_required(login_url='login')
def dashboard(request):
    query_nama = request.GET.get('nama', '').strip()
    query_tanggal = request.GET.get('tanggal', '').strip()
    
    pasien_list = Pasien.objects.all().order_by('-tanggal_skrining', '-id')
    
    if query_nama:
        pasien_list = pasien_list.filter(nama__icontains=query_nama)
    if query_tanggal:
        pasien_list = pasien_list.filter(tanggal_skrining=query_tanggal)
        
    # Hitung statistik untuk grafik & card ringkasan (khusus admin/keseluruhan data)
    total_pasien = Pasien.objects.count()
    total_berisiko = Pasien.objects.filter(status_risiko__icontains='Berisiko').count()
    total_normal = Pasien.objects.filter(status_risiko__icontains='Normal').count()
        
    context = {
        'pasien_list': pasien_list,
        'query_nama': query_nama,
        'query_tanggal': query_tanggal,
        'total_pasien': total_pasien,
        'total_berisiko': total_berisiko,
        'total_normal': total_normal,
    }
    return render(request, 'dashboard.html', context)

@login_required(login_url='login')
def skrining_view(request):
    kriteria_gejala = Kriteria.objects.exclude(kode__startswith='PR').order_by('kode')

    if request.method == 'POST':
        data_profil = {
            'usia': int(request.POST.get('usia', 0)),
            'tinggi_badan': Decimal(request.POST.get('tinggi_badan', '0')),
            'berat_badan': Decimal(request.POST.get('berat_badan', '0')),
            'jarak_kehamilan': Decimal(request.POST.get('jarak_kehamilan', '0')),
            'jumlah_paritas': int(request.POST.get('jumlah_paritas', 0)),
        }
        
        # Tangkap tanggal input, jika kosong gunakan tanggal hari ini
        tanggal_input = request.POST.get('tanggal_skrining') or timezone.now().date()
        
        gejala_klinis_cf_user = {}
        for k in kriteria_gejala:
            nilai = request.POST.get(k.kode, '0.0')
            gejala_klinis_cf_user[k.kode] = Decimal(nilai)

        persentase, status, keterangan = hitung_cf(data_profil, gejala_klinis_cf_user)

        pasien = Pasien.objects.create(
            nama=request.POST.get('nama', 'Tanpa Nama'),
            usia=data_profil['usia'],
            tinggi_badan=data_profil['tinggi_badan'],
            berat_badan_trimester_1=data_profil['berat_badan'],
            jarak_kehamilan=data_profil['jarak_kehamilan'],
            jumlah_paritas=data_profil['jumlah_paritas'],
            tanggal_skrining=tanggal_input,
            persentase_risiko=persentase,
            status_risiko=status,
            keterangan_medis=keterangan 
        )
        
        return redirect('hasil_skrining', id=pasien.id)
    
    return render(request, 'skrining.html', {'kriteria_gejala': kriteria_gejala})

@login_required(login_url='login')
def hasil_skrining(request, id):
    pasien = get_object_or_404(Pasien, id=id)
    return render(request, 'detail_hasil.html', {'p': pasien})

@login_required(login_url='login')
def delete_pasien(request, id):
    pasien = get_object_or_404(Pasien, id=id)
    pasien.delete()
    return redirect('dashboard')

@login_required(login_url='login')
def export_excel(request):
    if not request.user.is_superuser:
        return redirect('dashboard')
        
    wb = Workbook()
    ws = wb.active
    ws.title = "Laporan Skrining"
    
    headers = ['No', 'Tanggal Kunjungan', 'Nama Pasien', 'Usia', 'Persentase Risiko', 'Status Risiko', 'Keterangan Medis']
    ws.append(headers)
    
    # Tangkap filter rentang tanggal dari request GET
    start_date = request.GET.get('start_date')
    end_date = request.GET.get('end_date')
    
    pasien_list = Pasien.objects.all().order_by('tanggal_skrining')
    if start_date and end_date:
        pasien_list = pasien_list.filter(tanggal_skrining__range=[start_date, end_date])
    elif start_date:
        pasien_list = pasien_list.filter(tanggal_skrining__gte=start_date)
    elif end_date:
        pasien_list = pasien_list.filter(tanggal_skrining__lte=end_date)
        
    for index, p in enumerate(pasien_list, start=1):
        ws.append([
            index,
            str(p.tanggal_skrining),
            p.nama,
            p.usia,
            f"{p.persentase_risiko}%",
            p.status_risiko,
            p.keterangan_medis
        ])
        
    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = 'attachment; filename="Laporan_Skrining_Puskesmas_Bihbul.xlsx"'
    wb.save(response)
    return response

@login_required(login_url='login')
def export_pdf(request):
    if not request.user.is_superuser:
        return redirect('dashboard')
        
    response = HttpResponse(content_type='application/pdf')
    response['Content-Disposition'] = 'attachment; filename="Laporan_Skrining_Puskesmas_Bihbul.pdf"'
    
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    elements = []
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'TitleStyle',
        parent=styles['Heading1'],
        fontSize=16,
        alignment=1,
        textColor=colors.HexColor('#d81b60')
    )
    
    elements.append(Paragraph("<b>PUSKESMAS BIHBUL</b>", title_style))
    elements.append(Paragraph("<b>Laporan Rekapitulasi Skrining Risiko Kehamilan</b>", ParagraphStyle('Sub', parent=styles['Normal'], alignment=1, fontSize=12, textColor=colors.gray)))
    elements.append(Spacer(1, 15))
    
    # Tangkap filter rentang tanggal untuk ditampilkan di PDF jika ada
    start_date = request.GET.get('start_date')
    end_date = request.GET.get('end_date')
    
    info_periode = f"Periode: {start_date or 'Awal'} s/d {end_date or 'Sekarang'}" if (start_date or end_date) else "Semua Periode"
    elements.append(Paragraph(f"<i>{info_periode}</i>", ParagraphStyle('Periode', parent=styles['Normal'], alignment=1, fontSize=10, textColor=colors.darkgray)))
    elements.append(Spacer(1, 15))
    
    data = [['No', 'Tanggal', 'Nama Pasien', 'Usia', 'Status', 'Persentase']]
    
    pasien_list = Pasien.objects.all().order_by('tanggal_skrining')
    if start_date and end_date:
        pasien_list = pasien_list.filter(tanggal_skrining__range=[start_date, end_date])
    elif start_date:
        pasien_list = pasien_list.filter(tanggal_skrining__gte=start_date)
    elif end_date:
        pasien_list = pasien_list.filter(tanggal_skrining__lte=end_date)
        
    for index, p in enumerate(pasien_list, start=1):
        data.append([
            str(index),
            str(p.tanggal_skrining),
            p.nama,
            f"{p.usia} thn",
            p.status_risiko,
            f"{p.persentase_risiko}%"
        ])
        
    t = Table(data, colWidths=[30, 70, 150, 50, 80, 70])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#ff69b4')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0,0), (-1,0), 8),
        ('BACKGROUND', (0,1), (-1,-1), colors.HexColor('#fff9fa')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#f8bbd0')),
        ('FONTSIZE', (0,0), (-1,-1), 9),
    ]))
    
    elements.append(t)
    doc.build(elements)
    
    pdf = buffer.getvalue()
    buffer.close()
    response.write(pdf)
    return response