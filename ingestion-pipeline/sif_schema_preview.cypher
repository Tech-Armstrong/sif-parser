// Draws the SIF graph design in Neo4j Browser using virtual nodes (APOC): writes nothing to the database.
// db.schema.visualization() only shows relationships once real data exists.
CALL apoc.create.vNode(['SIF_FUND'], {name: 'SIF_FUND', properties: 'sifCode, sifFundName, sifCategory'}) YIELD node AS fund
CALL apoc.create.vNode(['SIF_MONTHLY_SNAPSHOT'], {name: 'SIF_MONTHLY_SNAPSHOT', properties: 'sifCode, sifYear, sifMonth, sifReportedDate, sifAUM'}) YIELD node AS snapshot
CALL apoc.create.vNode(['SIF_ASSET_CLASS'], {name: 'SIF_ASSET_CLASS', properties: 'sifName, sifSnapshotYear, sifSnapshotMonth'}) YIELD node AS assetClass
CALL apoc.create.vNode(['SIF_INSTRUMENT'], {name: 'SIF_INSTRUMENT', properties: 'sifISIN, sifName, sifAssetClass, sifSubType, sifSector, sifIndustry, sifSymbol, sifMarketCap'}) YIELD node AS instrument
CALL apoc.create.vNode(['SIF_FUND_MANAGER'], {name: 'SIF_FUND_MANAGER', properties: 'sifName'}) YIELD node AS manager
CALL apoc.create.vRelationship(manager, 'SIF_MANAGES', {properties: 'sifStartDate, sifEndDate'}, fund) YIELD rel AS manages
CALL apoc.create.vRelationship(fund, 'SIF_LATEST_MONTHLY_SNAPSHOT', {}, snapshot) YIELD rel AS latest
CALL apoc.create.vRelationship(snapshot, 'SIF_PREVIOUS_MONTHLY_SNAPSHOT', {}, snapshot) YIELD rel AS previous
CALL apoc.create.vRelationship(snapshot, 'SIF_HAS_ASSET_CLASS', {}, assetClass) YIELD rel AS hasClass
CALL apoc.create.vRelationship(assetClass, 'SIF_HAS_INSTRUMENT', {properties: 'sifHoldings, sifSubType, sifIncludedInNetAssets'}, instrument) YIELD rel AS holds
RETURN fund, snapshot, assetClass, instrument, manager, manages, latest, previous, hasClass, holds;
