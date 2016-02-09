# instagetter
getting pictures from instagramm and saving them locally

SHIT they changed API, added that  ig_cache_key!! only for JPG, not for MP4!!! 
only one script is fixed for now  -  instagetall.vbs

  "standard_resolution": {
                    "url": "https://scontent.cdninstagram.com/t51.2885-15/s640x640/sh0.08/e35/12568721_722372717863385_245242477_n.jpg?ig_cache_key=MTE3MTk3Njg5MDg2ODY0OTgyMw%3D%3D.2.l",
                    "width": 640,
                    "height": 640
                }
            },
            


Since it is VBScript, it works in any Windows (no it is JScript.. (because of JSON) which call VBscript.. sometimes.. because JScript does not have popup dialog boxes... and that shit DOES NOT WORK ON WIN7,  so it is XP only... sad..)

(probably i need to keep it as VBScript.. with ugly strings matching for JSON.. but it will work on Win7-8)


 takes latest picture from instagram user feed

  first - get user id by name
  second - get JSON with links
  third - fing pic urg and save it locally

   client_id can be received by registration process on instagramm
   i would say MUST be received... do not reuse my id!!! (lol)


<strike>Since Imstagram sends JSON  it would be great to rewrite it in JScript!
 to do nice responce parcing  </strike>


Instagram started to store HiRes pictures, starting to change code for that 11/2/2015
