#!/usr/bin/env python
from debe import *

class CongestionLocating(object):
    def __init__(self):
        self.main()

    def utc_to_local(self, utc_dt):
        import pytz, time

        local_tz = pytz.timezone('Asia/Jakarta')
        local_dt = utc_dt.replace(tzinfo=pytz.utc).astimezone(local_tz)
        return local_tz.normalize(local_dt)

    def get_start_time(self):
        import time, timeit

        self.startText = time.strftime("%H:%M:%S")
        self.startTime = timeit.default_timer()

    def get_finish_time(self):
        import time, timeit, math

        self.elapsed = math.ceil(timeit.default_timer() - self.startTime)
        self.finishText = time.strftime("%H:%M:%S")

    # Get unchunked kinds
    def get_kinds_unchunked(self, limitQuery):
        data = []
        query = sessionPostgresTraffic.query(ProcessChunking).\
            filter(ProcessChunking.classification_id == 1, ProcessChunking.kind_processed == True, ProcessChunking.kind_chunked == False).\
            order_by(desc(ProcessChunking.t_time)).\
            limit(limitQuery)
        for q in query:
            data.append([q.raw_id, q.t_user_id])
        return data

    # Get words
    def get_word(self, datum):
        data = []
        query = sessionPostgresTraffic.query(ProcessLocating).\
            filter(ProcessLocating.raw_id == datum).\
            order_by(ProcessLocating.sequence).\
            all()

        for q in query:
            data.append([q.name, q.tag_name])
        return data

    def chunking(self, t_user_id, datum):
        from nltk import RegexpParser
        from grammars import grammars

        data = []
        ret = []
        for d in datum:
            ret.append((d[0], d[1]))
        gm = ''
        if t_user_id in grammars:
            for g in grammars[t_user_id]:
                gm += 'INFO:\n' + g[0] + '\n' + g[1] + '\n'
        else:
            for g in grammars[0]:
                gm += "INFO:\n" + g[0] + "\n" + g[1] + "\n"

        chunkinWithGrammar = RegexpParser(gm)
        chunkResult = chunkinWithGrammar.parse(ret)
        data.append(chunkResult)
        return data

    def find_location_condition(self, datum):
        from nltk import tree
        ret = []
        for subtree in datum[0].subtrees(filter = lambda t: t.label() == 'INFO'):
            ret.append(subtree.leaves())

        # rearrange list to [place, condition]
        results = []
        placeTemp = ''
        conditionTemp = ''

        for words in ret:
            isJJExist = False
            for w in words:
                if w[1] != 'JJ':
                    placeTemp += ' ' + w[0]
                elif w[1] == 'JJ':
                    conditionTemp = w[0]
                    isJJExist = True
            if isJJExist:
                # Remove space before and after
                placeTemp = placeTemp.strip()
                results.append([placeTemp, conditionTemp])
                # Reset temporary variable
                placeTemp = ''
                conditionTemp = ''

        return results

    def update_kind_data(self, raw_id):
        query = sessionPostgresTraffic.query(Kind).\
            filter(Kind.raw_id == raw_id).\
            first()
        query.chunked = True
        sessionPostgresTraffic.commit()

    def update_word_data(self, raw_id):
        query = sessionPostgresTraffic.query(Word).\
            filter(Word.raw_id == raw_id).\
            all()
        for q in query:
            q.processed = True
        sessionPostgresTraffic.commit()

    def insert_chunk_data(self, data):
        for d in data[4]:
            temp = Chunk(
                raw_id = data[0],
                place = d[0],
                condition = d[1],
                weather = 'cerah'
            )
            sessionPostgresTraffic.add(temp)

    def main(self):
        limitQuery = 50
        results = []
        data = self.get_kinds_unchunked(limitQuery)
        if len(data) > 0:
            for d in data:
                dWord = self.get_word(d[0])
                dChunk = self.chunking(d[1], dWord)
                dLocating = self.find_location_condition(dChunk)
                self.insert_chunk_data([d[0], d[1], dWord, dChunk, dLocating])
                results.append([d[0], d[1], dWord, dChunk, dLocating])
            sessionPostgresTraffic.commit()
        else:
            results = []

        if len(results) > 0:
            for r in results:
                self.update_kind_data(r[0])
                self.update_word_data(r[0])
                print(r)

def main():
    CongestionLocating()

if __name__ == '__main__':
    main()
