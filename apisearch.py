# Librería necesaria
import requests
  
# Definiendo el endpoint
API_ENDPOINT = "http://api.arrow.com/itemservice/v4/en/search/token?login=brunosenzio&apikey=3ce2e3d8b027928abc0256adecf48ab98f6c3858b69d9ed6c3f0bfd49f6a3347&search_token=BAV99&rows=10"
  
# Haciendo solicitud y obteniendo respuesta
r = requests.post(url = API_ENDPOINT)
  
# extrayendo e imprimiendo respuesta
respuesta = r.text
print("Respuesta:%s"%respuesta)