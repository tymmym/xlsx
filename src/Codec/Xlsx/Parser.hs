{-# LANGUAGE NoMonomorphismRestriction #-}
{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE TupleSections #-}

module Codec.Xlsx.Parser(
  xlsx,
  sheet,
  cellSource,
  sheetRowSource
  ) where

import           Control.Applicative
import           Control.Monad (join)
import           Control.Monad.IO.Class()
import           Data.Function (on)
import qualified Data.IntMap as M
import qualified Data.IntSet as S
import           Data.List
import qualified Data.Map as Map
import           Data.Maybe
import           Data.Ord
import           Prelude hiding (sequence)

import           Data.Text (Text)
import qualified Data.Text as T
import qualified Data.Text.Read as T
import qualified Data.ByteString.Lazy as L
import           Data.ByteString.Lazy.Char8()

import           Codec.Archive.Zip
import           Data.Conduit
import qualified Data.Conduit.List as CL
import qualified Data.Conduit.Util as CU
import           Data.XML.Types
import           System.FilePath
import           Text.XML as X
import           Text.XML.Cursor
import qualified Text.XML.Stream.Parse as Xml

import           Codec.Xlsx


type MapRow = Map.Map Text Text


-- | Preload 'Xlsx' fields
xlsx :: FilePath -> IO Xlsx
xlsx fname = withArchive fname $
  Xlsx fname <$> getSharedStrings <*> getStyles <*> getWorksheetFiles


-- | Get data from specified worksheet as conduit source.
cellSource :: Xlsx -> Int -> [Text] -> Sink [Cell] (ResourceT Archive) a -> IO a
cellSource x sheetN cols sink = withArchive (xlArchive x) $
  getSheetCells x sheetN  $ filterColumns (S.fromList $ map col2int cols)
                         =$ groupRows
                         =$ reverseRows
                         =$ sink


decimal :: Monad m => Text -> m Int
decimal t = case T.decimal t of
  Right (d, _) -> return d
  _ -> fail "invalid decimal"

rational :: Monad m => Text -> m Double
rational t = case T.rational t of
  Right (r, _) -> return r
  _ -> fail "invalid rational"


sheet :: Xlsx -> Int -> IO Worksheet
sheet Xlsx{xlArchive=ar, xlSharedStrings=ss, xlWorksheetFiles=sheets} sheetN
  | sheetN < 0 || sheetN >= length sheets
    = fail "parseSheet: Invalid sheet number"
  | otherwise
    = collect <$> parse
  where
    filename = wfPath $ sheets !! sheetN
    sName = wfName $ sheets !! sheetN
    file = L.fromChunks <$> (withArchive ar $ sourceEntry filename CL.consume)  -- FIXME: fromChunks
    doc = do
      f <- file
      case parseLBS def f of
        Left _ -> error "could not read file"
        Right d -> return d
    tc :: IO Cursor
    tc = fromDocument <$> doc
    parse = tc >>= (\x -> return (x $/ parseColumns, x $/ parseRows))
    parseColumns :: Cursor -> [ColumnsWidth]
    parseColumns = element (n"cols") &/ element (n"col") >=> parseColumn
    parseColumn :: Cursor -> [ColumnsWidth]
    parseColumn c = do
      min <- c $| attribute "min" >=> decimal
      max <- c $| attribute "max" >=> decimal
      width <- c $| attribute "width" >=> rational
      return $ ColumnsWidth min max width
    parseRows :: Cursor -> [(Int, Maybe Double, [(Int, Int, CellData)])]
    parseRows = element (n"sheetData") &/ element (n"row") >=> parseRow
    parseRow c = do
      r <- c $| attribute "r" >=> decimal
      let ht = if attribute "customHeight" c == ["true"] 
               then listToMaybe $ c $| attribute "ht" >=> rational
               else Nothing
      return (r, ht, c $/ element (n"c") >=> parseCell)
    parseCell :: Cursor -> [(Int, Int, CellData)]
    parseCell cell = do
      (c, r) <- T.span (>'9') <$> (cell $| attribute "r")
      return (col2int c, int r, CellData s d)
      where
        s = listToMaybe $ cell $| attribute "s" >=> decimal
        t = fromMaybe "n" $ listToMaybe $ cell $| attribute "t"
        d = listToMaybe $ cell $/ element (n"v") &/ content >=> extractValue
        extractValue v = case t of
          "n" ->
            case T.rational v of
              Right (d, _) -> [CellDouble d]
              _ -> []
          "s" ->
            case T.decimal v of
              Right (d, _) -> maybeToList $ fmap CellText $ M.lookup d ss
              _ -> []
          _ -> []
    collect (cw, rd) = Worksheet sName minX maxX minY maxY cw rowMap cellMap
      where
        (rowMap, (minX, maxX, minY, maxY, cellMap)) = foldr collectRow rInit rd
        rInit = (Map.empty, (maxBound, minBound, maxBound, minBound, Map.empty))
        collectRow (_, Nothing, cells) (rowMap, cellData) = 
          (rowMap, foldr collectCell cellData cells)
        collectRow (n, Just h, cells) (rowMap, cellData) = 
          (Map.insert n h rowMap, foldr collectCell cellData cells)
        collectCell (x, y, cd) (minX, maxX, minY, maxY, cellMap) =
          (min minX x, max maxX x, min minY y, max maxY y, Map.insert (x,y) cd cellMap)
    

-- | Get all rows from specified worksheet.
sheetRowSource :: Xlsx -> Int -> Sink MapRow (ResourceT Archive) a -> IO a
sheetRowSource x sheetN sink = withArchive (xlArchive x) $
  getSheetCells x sheetN  $ groupRows
                         =$ reverseRows
                         =$ mkMapRows
                         =$ sink

-- | Make 'Conduit' from 'mkMapRowsSink'.
mkMapRows :: Monad m => Conduit [Cell] m MapRow
mkMapRows = CL.sequence mkMapRowsSink =$= CL.concatMap id


-- | Make 'MapRow' from list of 'Cell's.
mkMapRowsSink :: Monad m => Consumer [Cell] m [MapRow]
mkMapRowsSink = do
    header <- fromMaybe [] <$> CL.head
    rows   <- CL.consume

    return $ map (mkMapRow header) rows
  where
    mkMapRow header row = Map.fromList $ zipCells header row

    zipCells :: [Cell] -> [Cell] -> [(Text, Text)]
    zipCells []            _          = []
    zipCells header        []         = map (\h -> (txt h, "")) header
    zipCells header@(h:hs) row@(r:rs) =
        case comparing (fst . cellIx) h r of
          LT -> (txt h , ""   ) : zipCells hs     row
          EQ -> (txt h , txt r) : zipCells hs     rs
          GT -> (""    , txt r) : zipCells header rs

    txt = fromMaybe "" . cv
    cv Cell{cellData=CellData{cdValue=Just(CellText t)}} = Just t
    cv _ = Nothing

reverseRows :: Monad m => Conduit [a] m [a]
reverseRows = CL.map reverse
groupRows = CL.groupBy ((==) `on` (snd.cellIx))
filterColumns cs = CL.filter ((`S.member` cs) . col2int . fst . cellIx)


getSheetCells :: Xlsx -> Int -> Sink Cell (ResourceT Archive) a -> Archive a
getSheetCells (Xlsx{xlArchive=ar, xlSharedStrings=ss, xlWorksheetFiles=sheets}) sheetN sink
  | sheetN < 0 || sheetN >= length sheets
    = error "parseSheet: Invalid sheet number"
  | otherwise
    = xmlSource (wfPath $ sheets !! sheetN) $ mkXmlCond (getCell ss) =$ sink


-- | Parse single cell from xml stream.
getCell :: MonadThrow m => M.IntMap Text -> Consumer Event m (Maybe Cell)
getCell ss = Xml.tagName (n"c") cAttrs cParser
  where
    cAttrs = do
      cellIx  <- Xml.requireAttr  "r"
      style   <- Xml.optionalAttr "s"
      typ <- Xml.optionalAttr "t"
      Xml.ignoreAttrs
      return (cellIx,style,typ)

    maybeCellDouble Nothing = Nothing
    maybeCellDouble (Just t) = either (const Nothing) (\(d,_) -> Just (CellDouble d)) $ T.rational t

    cParser (ix,style,typ) = do
      val <- case typ of
          Just "inlineStr" -> liftA (fmap CellText) (tagSeq ["is", "t"])
          Just "s" -> liftA (fmap CellText) (tagSeq ["v"] >>=
                                             return . join . fmap ((`M.lookup` ss).int))
          Just "n" -> liftA maybeCellDouble $ tagSeq ["v"]
          _        -> liftA maybeCellDouble $ tagSeq ["v"]
      return $ Cell (mkCellIx ix) $ CellData (int <$> style) val

    mkCellIx ix = let (c,r) = T.span (>'9') ix
                  in (c,int r)


-- | Add sml namespace to name
n x = Name
  {nameLocalName = x
  ,nameNamespace = Just "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  ,namePrefix = Nothing}

-- | Add office document relationship namespace to name
odr x = Name
  {nameLocalName = x
  ,nameNamespace = Just "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  ,namePrefix = Nothing}

-- | Add package relationship namespace to name
pr x = Name
  {nameLocalName = x
  ,nameNamespace = Just "http://schemas.openxmlformats.org/package/2006/relationships"
  ,namePrefix = Nothing}


-- | Get text from several nested tags
tagSeq :: MonadThrow m => [Text] -> Consumer Event m (Maybe Text)
tagSeq (x:xs)
  = Xml.tagNoAttr (n x)
  $ foldr (\x -> Xml.force "" . Xml.tagNoAttr (n x)) Xml.content xs

tagSeq _ = error "no tags in tag sequence"


-- | Stream xml events from the specified file inside the zip archive.
xmlSource :: FilePath -> Sink Event (ResourceT Archive) a -> Archive a
xmlSource fname sink
  = sourceEntry fname $ Xml.parseBytes Xml.def =$ sink


-- Get shared strings (if there are some) into IntMap.
getSharedStrings :: Archive (M.IntMap Text)
getSharedStrings
  = (M.fromAscList . zip [0..]) <$> xmlSource "xl/sharedStrings.xml" sinkText

-- | Fetch all text from xml stream.
sinkText = mkXmlCond Xml.contentMaybe =$ CL.consume


getStyles :: Archive Styles
getStyles
  = Styles . L.fromChunks <$> sourceEntry "xl/styles.xml" CL.consume  -- FIXME: fromChunks only in bytestring >= 0.10.*

getWorksheetFiles :: Archive [WorksheetFile]
getWorksheetFiles = do
  sheetData <- xmlSource "xl/workbook.xml" $ mkXmlCond getSheetData =$ CL.consume
  wbRels <- getWbRels
  return [WorksheetFile n ("xl" </> T.unpack (fromJust $ lookup rId wbRels)) | (n, rId) <- sheetData]

getSheetData = Xml.tagName (n"sheet") attrs return
  where
    attrs = do
      name <- Xml.requireAttr "name"
      rId  <- Xml.requireAttr (odr "id")
      Xml.ignoreAttrs
      return (name, rId)

getWbRels :: Archive [(Text, Text)]
getWbRels
  = xmlSource "xl/_rels/workbook.xml.rels" parseWbRels


parseWbRels = Xml.force "relationships required" $
              Xml.tagNoAttr (pr"Relationships") $
              Xml.many $ Xml.tagName (pr"Relationship") attr return
  where
    attr = do
      target <- Xml.requireAttr "Target"
      id <- Xml.requireAttr "Id"
      Xml.ignoreAttrs
      return (id, target)

---------------------------------------------------------------------


int :: Text -> Int
int = either error fst . T.decimal


-- | Create conduit from xml sink
-- Resulting conduit filters nodes that `f` can consume and skips everything
-- else.
--
-- FIXME: Some benchmarking required: maybe it's not very efficient to `peek`i
-- each element twice. It's possible to swap call to `f` and `CL.peek`.
mkXmlCond f = self
  where
    self = CL.peek >>= maybe           -- try get current event form the stream
           (return ())                 -- stop if stream is empty
           (\_ -> yieldEvent >> self)  -- yield event and loop
    yieldEvent = f >>= maybe           -- try consume current event
                 (CL.drop 1)           -- skip it if can't process
                 yield                 -- yield result otherwise

-- mkXmlCond f = self
--   where
--     self = f >>= maybe                 -- try consume current event
--            (CL.drop 1 >> peekEvent)    -- skip it if can't process, then try to peek next event
--            (\e -> yield e >> self)     -- yield result otherwise and loop
--     peekEvent = CL.peek >>= maybe      -- try get current event form the stream
--                 (return ())            -- stop if stream is empty
--                 (const self)           -- loop otherwise
