// app/api/send-whatsapp/route.ts
import { NextResponse } from 'next/server';

export async function POST(request: Request) {
  try {
    const { mensaje, numero } = await request.json();
    
    // Limpiar el número (solo dígitos)
    const numeroLimpio = numero.replace(/[^0-9]/g, '');
    
    // Tu API Key de CallMeBot (es tu número de teléfono con código de país)
    // Ejemplo: Si tu número es +57 3004455245, la API Key es 573004455245
    const API_KEY = '573004455245'; // <--- REEMPLAZA CON TU NÚMERO
    
    // Codificar mensaje para URL
    const mensajeCodificado = encodeURIComponent(mensaje);
    
    // URL de CallMeBot
    const url = `https://api.callmebot.com/whatsapp.php?phone=${numeroLimpio}&text=${mensajeCodificado}&apikey=${API_KEY}`;
    
    console.log('📤 Enviando mensaje a:', numeroLimpio);
    console.log('📝 Mensaje:', mensaje.substring(0, 100) + '...');
    
    // Enviar mensaje
    const response = await fetch(url);
    const resultado = await response.text();
    
    console.log('📥 Respuesta:', resultado);
    
    // Verificar si fue exitoso
    if (resultado.includes('OK') || resultado.includes('Message') || resultado === 'true') {
      return NextResponse.json({ 
        success: true, 
        message: 'Mensaje enviado correctamente' 
      });
    } else {
      return NextResponse.json({ 
        success: false, 
        error: resultado 
      }, { status: 500 });
    }
    
  } catch (error) {
    console.error('❌ Error:', error);
    return NextResponse.json({ 
      success: false, 
      error: String(error) 
    }, { status: 500 });
  }
}

// Endpoint GET para probar
export async function GET() {
  return NextResponse.json({ 
    status: 'API de WhatsApp funcionando ✅',
    instrucciones: 'Envía un POST a este endpoint con { mensaje, numero }'
  });
}