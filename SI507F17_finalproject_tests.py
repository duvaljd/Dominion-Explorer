import unittest
from SI507F17_finalproject import *

class test_cacheFiles(unittest.TestCase):
    def setUp(self):
        self.cache = open(CACHE_FNAME, "r")
        self.data = json.loads(self.cache.read())

    def test_cacheExists(self):
        self.assertIsNotNone(self.cache.read(), "No cached data.")

    def test_cacheLinks(self):
        self.assertTrue("http://wiki.dominionstrategy.com/index.php/sets setlist" in self.data, "List of sets is not in cached data.")
        self.assertTrue("http://wiki.dominionstrategy.com/index.php/list_of_cards card" in self.data, "List of cards is not in cached data.")

    def tearDown(self):
        self.cache.close()

class test_databaseRequests(unittest.TestCase):
    def setUp(self):
        self.countCards = len(fetchAll_cardIDs())
        self.countRecs = len(fetchAll_recIDs())
        self.countSets = len(fetchAll_setIDs())

    def test_numberOf_cardResults(self):
        self.assertTrue(self.countCards == 371, "Fetch cards function got {} instead of 371".format(self.countCards))

    def test_numberOf_recResults(self):
        self.assertTrue(self.countRecs == 158, "Fetch recs function got {} instead of 158".format(self.countRecs))

    def test_numberOf_setResults(self):
        self.assertTrue(self.countSets == 12, "Fetch sets function got {} instead of 12".format(self.countSets))


class test_classes(unittest.TestCase):
    def setUp(self):
        self.cardList = [7, 23, 43, 89, 109, 139, 193, 227, 251, 324]
        self.cardNames = ["Village", "Council Room", "Secret Chamber", "Treasure Map", "Golem", "Fortune Teller", "Storeroom", "Dame Natalie", "Page", "Bard"]

        self.setList = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
        self.setNames = ["Dominion", "Intrigue", "Seaside", "Alchemy", "Prosperity", "Cornucopia", "Hinterlands", "Dark Ages", "Guilds", "Adventures", "Empires", "Nocturne"]

    def test_initCard(self):
        n = 0
        for card in self.cardList:
            self.assertTrue(Card(self.cardList[n]).cardName == (self.cardNames[n]), "card.cardName class variable incorrect.")
            n += 1

    def test_initSet(self):
        n = 0
        for cset in self.setList:
            self.assertTrue(Set(self.setList[n]).setName == self.setNames[n], "set.setName variable incorrect.")
            n += 1

    def test_Set_contains(self):
        self.assertTrue("Cellar" in Set(1), "Contains method did not find Cellar in Dominion.")
        self.assertFalse("Sacred Grove" in Set(11), "Contains method found Sacred Grove in Empires.")

class test_excelFiles(unittest.TestCase):
    def setUp(self):
        self.cardList = []
        
        for x in range(10):
            rn = random.randrange(0, 371)
            self.cardList.append(Card(rn))

        for card in self.cardList:
            makeFile_card(card.cid)

        for card in self.cardList:
            makeFile_set(card.cid)

    def test_makeList(self):
        makeList()
        with open('dominion.xlsx', 'r', encoding='latin-1') as f:
            self.assertTrue(f.read(), "Did not find 'dominion.xlsx'.")


    def test_makeFile_card(self):
        for card in self.cardList:
            with open("card_{}.xlsx".format(card.cardName.strip().replace('\'', '').lower()), 'r', encoding='latin-1') as f:
                self.assertTrue(f.read(), "Did not find file for random card.")

    def test_makeFile_set(self):
        for card in self.cardList:
            with open("set_{}.xlsx".format(card.setName.strip().replace('\'', '').lower()), 'r', encoding='latin-1') as f:
                self.assertTrue(f.read(), "Did not find file for random set.")

    def test_makeFile_recs(self):
        for card in self.cardList:
            if makeFile_recs(card.cid) == True:
                with open("recs_{}.xlsx".format(card.cardName.strip().replace('\'', '').lower()), 'r', encoding='latin-1') as f:
                    self.assertTrue(f.read(), "Did not find file for random set.")

if __name__ == "__main__":
    unittest.main(verbosity=2)