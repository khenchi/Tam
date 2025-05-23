Voici une version corrigée et améliorée de ton texte, avec une meilleure formulation, orthographe et clarté technique :


---

Objet : Point sur la configuration Oracle et l’accès en VBA

Bonjour,

Suite à mes échanges avec Yassine, voici comment cela fonctionne pour le client Oracle :

Définir la variable d’environnement TNS_ADMIN afin qu’elle pointe vers le bon répertoire (contenant tnsnames.ora).

Modifier la variable d’environnement Path en supprimant les références à Oracle en version 32 bits.

Adapter la méthode de connexion, en VBA, à la base de données Oracle : utiliser OLE DB (Provider=) au lieu de ODBC (Driver=).


Concernant l'utilisation de la bibliothèque AdfinX Real Time 6.0 Library, je ne trouve pas d’équivalent compatible avec Excel/VBA en 64 bits.
A priori, cette bibliothèque ne fonctionne qu’en environnement 32 bits.

Bien à vous,


---

Souhaites-tu que je l’adapte à un ton plus formel ou plus collaboratif selon le destinataire ?





If your database isn’t defined in the tnsnames.ora file, you have two options:


---

Option 1: Use a full connect descriptor directly (no tnsnames.ora needed)

This is what you're already doing with a connection string like this:

Conn.Open "Provider=OraOLEDB.Oracle;" & _
          "Data Source=(DESCRIPTION=" & _
              "(ADDRESS=(PROTOCOL=TCP)(HOST=histore.int.cmr.emea.cib)(PORT=1690))" & _
              "(CONNECT_DATA=(SERVICE_NAME=HISTOPE.WORLD)));" & _
          "User ID=HISTO_RM;" & _
          "Password=GHJDJEffje375_;"

This approach is called a "fully qualified connect descriptor" and completely bypasses the need for tnsnames.ora.

> Advantage: You don’t rely on TNS files or environment configuration
Disadvantage: Harder to maintain or share — full host and service info must be stored in your code




---

Option 2: Define the database in your tnsnames.ora file (recommended if reused)

Create or update the file tnsnames.ora, usually found (or placed) in:

%TNS_ADMIN%\tnsnames.ora

Add something like:

HISTOPE_ALIAS =
  (DESCRIPTION =
    (ADDRESS = (PROTOCOL = TCP)(HOST = histore.int.cmr.emea.cib)(PORT = 1690))
    (CONNECT_DATA = (SERVICE_NAME = HISTOPE.WORLD))
  )

Then in VBA:

Conn.Open "Provider=OraOLEDB.Oracle;Data Source=HISTOPE_ALIAS;User ID=HISTO_RM;Password=GHJDJEffje375_;"

> Advantage: Cleaner, easier to maintain, changes only in one place
Disadvantage: Requires proper TNS configuration and TNS_ADMIN pointing to the file




---

Summary

Approach	Needs tnsnames.ora?	Editable outside VBA?	Best for

Full connect descriptor	No	No	One-off or hardcoded use
TNS alias (via tnsnames.ora)	Yes	Yes	Shared/multi-user environments



---

Do you want help writing or checking your tnsnames.ora file? I can guide you.

