{-# LANGUAGE OverloadedStrings #-}

module Main (main) where

import qualified Data.Map as M
--import           Data.Time.Calendar
--import           Data.Time.LocalTime
import           System.FilePath

import           Data.Text (Text)
import           System.IO.Temp
import           Test.HUnit (Assertion, assertEqual)
import           Test.QuickCheck
import           Test.Framework (Test, defaultMain, testGroup)
import           Test.Framework.Providers.HUnit
import           Test.Framework.Providers.QuickCheck2 (testProperty)

import           Codec.Xlsx
import           Codec.Xlsx.Writer
import           Codec.Xlsx.Parser

main :: IO ()
main = defaultMain tests

tests :: [Test]
tests =
    [ testGroup "Behaves to spec"
        [ testProperty "col to name" prop_col2name
        , testCase     "write read " writeReadXlsx
        ]
    ]

prop_col2name :: Positive Int -> Bool
prop_col2name (Positive i) = i == (col2int $ int2col i)
  where types = (i::Int)

writeReadXlsx :: Assertion
writeReadXlsx = withSystemTempDirectory "xlsx" $ \dir -> do
  let path = dir </> "test.xlsx"
  writeXlsx path [sheetIn]
  x <- xlsx path
  sheetOut <- sheet x 0
  assertEqual "" (wsCells sheetIn) $ removeLastRow (wsCells sheetOut)  -- FIXME: writeXlsx adds additional blank row
  where
    sheetIn = fromList "List" cols rows cells
    cols    = [ColumnsWidth 1 10 15]
    rows    = M.fromList [(1,50)]
    cells   = [[xText "column1", xText "column2", xText "column3", xDouble 42.12345, xText  "False"]]  -- TODO: test with dates: (xDate $! LocalTime (fromGregorian 2012 05 06) (TimeOfDay 7 30 50))

removeLastRow :: (Ord c, Ord r) => M.Map (c, r) a -> M.Map (c, r) a
removeLastRow m =
  M.filterWithKey (\k _ -> not . (== lastRow) $ snd k) m
  where
    lastRow = snd . fst $ M.findMax m

xText :: Text -> Maybe CellData
xText t = Just CellData{cdValue=Just $ CellText t, cdStyle=Just 0}

xDouble :: Double -> Maybe CellData
xDouble d = Just CellData{cdValue=Just $ CellDouble d, cdStyle=Just 0}

-- xDate :: LocalTime -> Maybe CellData
-- xDate d = Just CellData{cdValue=Just $ CellLocalTime d, cdStyle=Just 0}
