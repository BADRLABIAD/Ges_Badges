-- =====================================================================
-- Elargissement de CODE_COMPETENCE.CODE_COMPETENCE (varchar(8) -> varchar(20))
-- Motif : les codes du referentiel achats (ex: ACH-BO-009) font 10 caracteres
-- et depassent la largeur actuelle -> erreur 8152 "String or binary data
-- would be truncated" a l'insertion.
--
-- POSTE_CRITERES.CODE_COMPETENCE (meme table logique, meme largeur 8) est
-- elargi en meme temps : sinon un code long accepte dans CODE_COMPETENCE
-- se ferait tout de meme tronquer/rejeter lors de l'insertion des liens
-- poste <-> competence.
--
-- Elargir un varchar est une operation non destructive (les valeurs
-- existantes, toutes <= 8 caracteres, restent inchangees). Aucun index/
-- containte ne doit etre supprime pour un simple agrandissement.
-- =====================================================================
SET NOCOUNT ON;
BEGIN TRANSACTION;

ALTER TABLE dbo.CODE_COMPETENCE ALTER COLUMN CODE_COMPETENCE varchar(20) NOT NULL;
ALTER TABLE dbo.POSTE_CRITERES  ALTER COLUMN CODE_COMPETENCE varchar(20) NOT NULL;

COMMIT TRANSACTION;

-- Verification post-execution :
-- SELECT TABLE_NAME, COLUMN_NAME, CHARACTER_MAXIMUM_LENGTH
-- FROM INFORMATION_SCHEMA.COLUMNS
-- WHERE COLUMN_NAME = 'CODE_COMPETENCE';
