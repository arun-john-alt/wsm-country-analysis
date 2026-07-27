"""
Localized expansion (READ-ONLY) for non-English markets. KP GenerateKeywordIdeas with the
locale's language constant + geo, seeded with DIALECT-ACCURATE hybrid seeds (from crawling
manageengine.com localized sites). -> kw_exp_ideas_raw_loc (country, product, theme, idea, sv...).
Generalizes kw_expand_de.py across FR/ES/MX/BR/IT/NL/PL/TR.
"""
import os, sys, uuid, time
sys.stdout.reconfigure(encoding='utf-8')
import bigboss_config as cfg
import google.auth
from google.cloud import bigquery
from google.ads.googleads.client import GoogleAdsClient
import pandas as pd
from datetime import date
bq = cfg.bq_client()
ADS_CONFIG = cfg.GOOGLE_ADS
CID=cfg.GOOGLE_ADS_CUSTOMER_ID; OUT=f"{cfg.BQ_PROJECT}.{cfg.BQ_DATASET}.kw_exp_ideas_raw_loc"; GEN=date.today().isoformat()

# locale -> (language_constant, geoTargetConstant, CampaignCountry)
META = {
 "fr": (1002, 2250, "France"), "es": (1003, 2724, "Spain"), "mx": (1003, 2484, "Mexico"),
 "br": (1014, 2076, "Brazil"), "it": (1004, 2380, "Italy"), "nl": (1010, 2528, "Netherlands"),
 "pl": (1030, 2616, "Poland"), "tr": (1037, 2792, "Turkey"),
}
# locale -> { (product, theme): [dialect seeds] }   (from localized-site crawls)
SEEDS = {
 "fr": {
  ("ADAP","Active Directory Audit"):["logiciel d'audit active directory","outil d'audit des modifications active directory","audit des modifications active directory temps réel","détecter qui a modifié un objet active directory","logiciel audit ad conformité rgpd"],
  ("ADAP","Account Lockout"):["analyse de verrouillage de compte active directory","outil verrouillage de compte active directory","trouver source verrouillage compte active directory"],
  ("ADAP","File Server Audit"):["solution de surveillance des serveurs de fichiers","audit serveur de fichiers windows","qui a supprimé un fichier serveur de fichiers"],
  ("ADMP","AD Management"):["logiciel de gestion active directory","outil de gestion des utilisateurs ad en masse","délégation administration active directory help desk","rapports active directory et microsoft 365","création automatique utilisateurs active directory"],
  ("ADSSP","SSPR"):["solution de réinitialisation de mot de passe en libre-service","réinitialisation mot de passe libre-service active directory","logiciel de déverrouillage de compte active directory","portail réinitialisation mot de passe active directory","réduire appels help desk mot de passe oublié"],
  ("ADSSP","MFA"):["authentification multifacteur mfa pour connexion windows","authentification unique sso entreprise active directory","mfa bureau à distance rdp","authentification multifacteur active directory sur site"],
  ("RMP","Microsoft 365 Backup"):["sauvegarde microsoft 365","sauvegarde microsoft 365 et exchange online","logiciel de sauvegarde microsoft 365","solution de sauvegarde office 365","sauvegarde office 365 tierce partie","sauvegarde sharepoint online"],
  ("RMP","Exchange Backup"):["logiciel de sauvegarde exchange","sauvegarde boîte aux lettres exchange","restaurer boîte aux lettres exchange"],
  ("RMP","Active Directory Backup"):["logiciel de sauvegarde active directory","outil de sauvegarde et restauration active directory","récupération objets active directory supprimés"],
  ("ELA","Log"):["logiciel de gestion des journaux siem","solution siem gestion des journaux","logiciel siem entreprise","outil de gestion des logs centralisé","détection des menaces siem"],
  ("ELA","Syslog"):["outil de gestion des logs windows et syslog","serveur syslog","logiciel serveur syslog windows"],
 },
 "es": {
  ("ADAP","Active Directory Audit"):["software auditoría active directory","herramienta auditoría cambios active directory tiempo real","auditoría servidor de archivos windows","monitorización usuarios privilegiados active directory"],
  ("ADAP","Account Lockout"):["análisis bloqueo de cuentas active directory","encontrar origen bloqueo de cuentas active directory"],
  ("ADMP","AD Management"):["software gestión active directory","herramienta gestión masiva usuarios active directory","delegación tareas help desk active directory","informes active directory office 365"],
  ("ADSSP","SSPR"):["software autoservicio restablecimiento contraseña active directory","desbloqueo de cuentas autoservicio usuarios","portal autoservicio restablecer contraseña active directory"],
  ("ADSSP","MFA"):["autenticación multifactor mfa active directory","inicio de sesión único sso empresarial","mfa para escritorio remoto rdp"],
  ("RMP","Microsoft 365 Backup"):["copia de seguridad microsoft 365","copia de seguridad exchange office 365","software copia de seguridad office 365","solución copia de seguridad microsoft 365","copia de seguridad sharepoint online"],
  ("RMP","Exchange Backup"):["copia de seguridad exchange","restaurar buzón exchange","software copia de seguridad exchange server"],
  ("RMP","Active Directory Backup"):["software copia de seguridad active directory","herramienta copia de seguridad y restauración active directory","recuperación objetos active directory eliminados"],
  ("ELA","Log"):["software gestión de registros siem","herramienta análisis registros eventos seguridad","solución siem empresarial","correlación eventos detección amenazas"],
  ("ELA","Syslog"):["servidor syslog software"],
 },
 "mx": {
  ("ADAP","Active Directory Audit"):["software de auditoría de active directory","herramienta de monitoreo de cambios en active directory","reportes de auditoría de active directory","monitoreo de inicio de sesión de usuarios"],
  ("ADAP","Account Lockout"):["análisis de bloqueo de cuentas active directory","encontrar origen de bloqueo de cuentas"],
  ("ADMP","AD Management"):["software de gestión de active directory","gestión masiva de usuarios active directory","reportes de active directory","delegación de tareas mesa de ayuda active directory"],
  ("ADSSP","SSPR"):["restablecimiento de contraseña de autoservicio active directory","resetear contraseña autoservicio","desbloqueo de cuenta de autoservicio","portal autoservicio de contraseñas"],
  ("ADSSP","MFA"):["autenticación multifactor mfa para active directory","inicio de sesión único sso active directory","mfa para escritorio remoto"],
  ("RMP","Microsoft 365 Backup"):["copia de seguridad de microsoft 365","respaldo de office 365","copia de seguridad de exchange online y microsoft 365","backup y restore de microsoft 365","software de respaldo de office 365","respaldo de sharepoint online"],
  ("RMP","Exchange Backup"):["respaldo de exchange","restaurar buzón de exchange","copia de seguridad de exchange server"],
  ("RMP","Active Directory Backup"):["software de copia de seguridad de active directory","respaldo y recuperación de active directory","recuperación de objetos eliminados de active directory"],
  ("ELA","Log"):["software de gestión de logs","herramienta de gestión de logs de seguridad","monitoreo de logs windows","solución siem","correlación de eventos de seguridad"],
  ("ELA","Syslog"):["gestión de logs syslog dispositivos de red"],
 },
 "br": {
  ("ADAP","Active Directory Audit"):["software de auditoria do active directory","ferramenta de monitoramento de mudanças no active directory","auditoria de quem alterou conta no active directory","relatórios de conformidade lgpd active directory"],
  ("ADAP","Logon/Logoff"):["auditoria de logon e logoff do active directory","monitoramento de logon de usuários windows"],
  ("ADAP","File Server Audit"):["monitoramento de integridade de arquivos servidor de arquivos","quem excluiu arquivo no servidor de arquivos"],
  ("ADMP","AD Management"):["ferramenta de gerenciamento do active directory","software de criação de usuários em massa active directory","relatórios do active directory exportar","delegação de help desk active directory"],
  ("ADSSP","SSPR"):["redefinição de senha de autoatendimento active directory","ferramenta de desbloqueio de conta de autoatendimento","portal de redefinição de senha sem help desk","redefinir senha active directory autoatendimento"],
  ("ADSSP","MFA"):["mfa para login do windows active directory","autenticação multifator active directory","mfa para área de trabalho remota rdp"],
  ("RMP","Microsoft 365 Backup"):["backup do microsoft 365","backup do microsoft 365 e restauração","software de backup do office 365","solução de backup do microsoft 365","backup do sharepoint online","backup de terceiros microsoft 365"],
  ("RMP","Exchange Backup"):["backup do exchange","backup e recuperação do exchange online","restaurar caixa de correio exchange"],
  ("RMP","Active Directory Backup"):["software de backup do active directory","ferramenta de restauração de objetos do active directory","recuperar objeto excluído active directory"],
  ("ELA","Log"):["software de gerenciamento de logs siem","análise de logs de eventos do windows","solução siem para empresas","correlação de eventos e detecção de ameaças"],
  ("ELA","Syslog"):["servidor syslog para windows e linux"],
 },
 "it": {
  ("ADAP","Active Directory Audit"):["software audit active directory","controllo modifiche active directory in tempo reale","audit file server chi ha modificato file","monitoraggio accessi active directory"],
  ("ADAP","Account Lockout"):["rilevare blocco account active directory","trovare origine blocco account active directory"],
  ("ADMP","AD Management"):["software gestione active directory","creazione utenti active directory in blocco","delega help desk active directory","report utenti active directory"],
  ("ADSSP","SSPR"):["reimpostazione self service password active directory","sblocco account self service utenti","portale reset password active directory"],
  ("ADSSP","MFA"):["mfa active directory autenticazione a più fattori","single sign on aziendale sso","mfa per desktop remoto rdp"],
  ("RMP","Microsoft 365 Backup"):["backup microsoft 365","backup e ripristino microsoft 365","software backup office 365","soluzione backup microsoft 365","backup sharepoint online"],
  ("RMP","Exchange Backup"):["backup exchange","backup exchange online","ripristino casella di posta exchange"],
  ("RMP","Active Directory Backup"):["software backup active directory","ripristino oggetti active directory eliminati","backup e ripristino active directory"],
  ("ELA","Log"):["software gestione log siem","analisi log eventi sicurezza","soluzione siem aziendale","correlazione eventi rilevamento minacce"],
  ("ELA","Syslog"):["monitoraggio log file server","server syslog"],
 },
 "nl": {
  ("ADAP","Active Directory Audit"):["active directory auditing software","active directory audit tool","real-time ad monitoring","compliance rapportage active directory","wie heeft gebruiker verwijderd active directory"],
  ("ADMP","AD Management"):["active directory beheer tool","ad management software","active directory rapportage software","ad gebruikers bulk beheren"],
  ("ADSSP","SSPR"):["self-service wachtwoord reset","wachtwoord opnieuw instellen tool","account ontgrendelen self-service","password self-service portal"],
  ("ADSSP","MFA"):["mfa active directory","single sign-on sso software","mfa voor remote desktop rdp"],
  ("RMP","Microsoft 365 Backup"):["office 365 back-up","microsoft 365 backup software","microsoft 365 back-up oplossing","office 365 backup solution","sharepoint online backup","exchange online back-up"],
  ("RMP","Exchange Backup"):["exchange back-up software","exchange mailbox back-up","exchange herstel tool"],
  ("RMP","Active Directory Backup"):["active directory back-up tool","ad back-up en herstel","verwijderde gebruiker herstellen active directory"],
  ("ELA","Log"):["log management software","siem software","siem tool","logboekbeheer software","windows event log monitoring"],
  ("ELA","Syslog"):["syslog monitoring tool","syslog server software"],
 },
 "pl": {
  ("ADAP","Active Directory Audit"):["audyt active directory","monitorowanie zmian active directory","oprogramowanie do audytu ad","raporty zgodności active directory"],
  ("ADAP","Account Lockout"):["analiza blokad kont active directory","źródło blokady konta active directory"],
  ("ADMP","AD Management"):["zarządzanie active directory","narzędzie do zarządzania active directory","zbiorcze zarządzanie użytkownikami ad","raportowanie active directory"],
  ("ADSSP","SSPR"):["samoobsługowe resetowanie hasła active directory","odblokowanie konta active directory","portal samoobsługowego resetowania hasła"],
  ("ADSSP","MFA"):["uwierzytelnianie wieloskładnikowe mfa active directory","jednokrotne logowanie sso","mfa do pulpitu zdalnego rdp"],
  ("RMP","Microsoft 365 Backup"):["kopia zapasowa microsoft 365","backup office 365","oprogramowanie do backupu microsoft 365","kopia zapasowa office 365","backup sharepoint online"],
  ("RMP","Exchange Backup"):["kopia zapasowa exchange","backup exchange online","odzyskiwanie skrzynki exchange"],
  ("RMP","Active Directory Backup"):["kopia zapasowa active directory","odzyskiwanie active directory","odzyskiwanie usuniętych obiektów active directory"],
  ("ELA","Log"):["zarządzanie logami siem","oprogramowanie siem","analiza logów windows","centralne zarządzanie logami"],
  ("ELA","Syslog"):["monitorowanie dzienników zdarzeń","serwer syslog"],
 },
 "tr": {
  ("ADAP","Active Directory Audit"):["active directory denetimi","gerçek zamanlı ad izleme","gpo değişiklik denetimi","active directory denetim yazılımı"],
  ("ADAP","Account Lockout"):["hesap kilitlenmesi çözümleyici","hesap kilitlenme kaynağı bulma active directory"],
  ("ADMP","AD Management"):["active directory yönetimi","ad raporlama aracı","toplu kullanıcı oluşturma","active directory yönetim yazılımı"],
  ("ADSSP","SSPR"):["self servis parola sıfırlama","hesap kilidini açma self servis","active directory parola sıfırlama aracı"],
  ("ADSSP","MFA"):["çok faktörlü kimlik doğrulama mfa active directory","tek oturum açma sso","uzak masaüstü rdp için mfa"],
  ("RMP","Microsoft 365 Backup"):["microsoft 365 yedekleme","office 365 yedekleme","microsoft 365 yedekleme yazılımı","office 365 yedekleme çözümü","sharepoint online yedekleme"],
  ("RMP","Exchange Backup"):["exchange yedekleme","exchange online yedekleme","exchange posta kutusu yedekleme"],
  ("RMP","Active Directory Backup"):["active directory yedekleme","ad yedekleme ve geri yükleme","silinen active directory nesnelerini kurtarma"],
  ("ELA","Log"):["günlük yönetimi siem","log yönetimi","olay günlüğü analizi","siem yazılımı kurumsal"],
  ("ELA","Syslog"):["syslog sunucusu"],
 },
}
ads = GoogleAdsClient.load_from_dict(ADS_CONFIG); svc = ads.get_service("KeywordPlanIdeaService")

def expand(country, lang, geo, product, theme, seeds):
    req=ads.get_type("GenerateKeywordIdeasRequest")
    req.customer_id=CID; req.language=f"languageConstants/{lang}"
    req.geo_target_constants.append(f"geoTargetConstants/{geo}")
    req.keyword_plan_network=ads.enums.KeywordPlanNetworkEnum.GOOGLE_SEARCH
    req.include_adult_keywords=False
    req.keyword_seed.keywords.extend(seeds)
    rows=[]; sset=set(s.lower().strip() for s in seeds)
    for a in range(3):
        try:
            resp=svc.generate_keyword_ideas(request=req)
            for r in resp:
                m=r.keyword_idea_metrics
                rows.append(dict(country=country,product=product,theme=theme,idea_keyword=r.text.lower().strip(),
                    avg_monthly_searches=int(m.avg_monthly_searches) if m.avg_monthly_searches else 0,
                    competition=m.competition.name if m.competition else "UNSPECIFIED",
                    bid_high=(m.high_top_of_page_bid_micros/1e6) if m.high_top_of_page_bid_micros else None,
                    is_seed=r.text.lower().strip() in sset, generated_date=GEN))
            return rows
        except Exception as e:
            print(f"   {country}/{product}/{theme} attempt {a+1}: {e}"); time.sleep(15*(a+1))
    return rows

allrows=[]
for loc,(lang,geo,country) in META.items():
    n=0
    for (product,theme),seeds in SEEDS[loc].items():
        ideas=expand(country,lang,geo,product,theme,seeds); allrows.extend(ideas); n+=len(ideas); time.sleep(0.6)
    print(f"[ok] {loc} {country:<14} ideas={n:,}")
df=pd.DataFrame(allrows)
print(f"\nTotal {len(df):,} ideas, unique {df['idea_keyword'].nunique():,}")
bq.load_table_from_dataframe(df,OUT,job_config=bigquery.LoadJobConfig(write_disposition="WRITE_TRUNCATE"),
    job_id_prefix=f"kwexploc_{uuid.uuid4().hex[:8]}_").result()
print(f"Wrote -> {OUT}")
