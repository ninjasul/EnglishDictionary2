namespace EnglishDictionary2
{
    using System;
    using System.Collections.Generic;

    internal class Vocabulary
    {
        private string                          name;                   // 단어 스펠링
        private string                          pronunciation;          // 발음기호
        private List<string>                    meanings;               // 뜻 리스트
        //private List<string> []                 sentencesArray;         // 예문 배열
        private List<List<string>>              sentencesList;          // 예문 리스트 리스트

        public Vocabulary()
        {
            this.name = string.Empty;
            this.pronunciation = string.Empty;
            this.meanings = null;
            //this.sentencesArray = null;
            this.sentencesList = null;
        }

        public Vocabulary(string name)
        {
            this.name = name;
            this.pronunciation = string.Empty;
            this.meanings = null;
            //this.sentencesArray = null;
            this.sentencesList = null;
        }

        /*
        public Vocabulary(string name, string pronunciation, List<string> meanings, List<string> [] sentencesArray)
        {
            this.name = name;
            this.pronunciation = pronunciation;
            this.meanings = meanings;
            this.sentencesArray = sentencesArray;
            this.sentencesList = null;
        }
        */

        public Vocabulary(string name, string pronunciation, List<string> meanings, List<List<string>> sentencesList)
        {
            this.name = name;
            this.pronunciation = pronunciation;
            this.meanings = meanings;
            this.sentencesList = sentencesList;
        }


        //public Vocabulary(List<string> meanings, List<string>[] sentencesArray)
        //{
        //    this.name = string.Empty;
        //    this.pronunciation = string.Empty;
        //    this.meanings = meanings;
        //    this.sentencesArray = sentencesArray;
        //}


        public Vocabulary(List<string> meanings, List<List<string>> sentencesList)
        {
            this.name = string.Empty;
            this.pronunciation = string.Empty;
            this.meanings = meanings;
            this.sentencesList = sentencesList;
        }


        public string NAME
        {
            get => 
                this.name;
            set => 
                this.name = value;
        }

        public string PRON
        {
            get => 
                this.pronunciation;
            set => 
                this.pronunciation = value;
        }

        public List<string> getMeanings()
        {
            return this.meanings;
        }

        public void setMeanings( List<string> meanings )
        {
            this.meanings = meanings;
        }

        /*
        public List<string> [] getSentencesArray()
        {
            return this.sentencesArray;
        }

        public void setSentencesArray(List<string>[] sentencesArray)
        {
            this.sentencesArray = sentencesArray;
        }
        */
        public List<List<string>> getSentencesList()
        {
            return this.sentencesList;
        }

        public void setSentencesList(List<List<string>> sentencesList)
        {
            this.sentencesList = sentencesList;
        }

        public string FRONT
        {
            get
            {
                string front = this.name + " " + this.pronunciation + Environment.NewLine;

                if (this.sentencesList.Count > 0)
                {
                    front += Environment.NewLine;
                    foreach (List<string> sentences in this.sentencesList)
                    {
                        foreach (string curSentence in sentences)
                        {
                            front += "- " + curSentence + Environment.NewLine;
                        }
                        front += Environment.NewLine;
                    }
                }

                front = StringUtil.removeRedundantNewLineCharacters(front);
                return front;
            }
        }

        public string BACK
        {
            get
            {
                string back = string.Empty;
                if(this.meanings.Count > 0)
                {
                    int i = 0;
                    foreach (string curMeaning in this.meanings)
                    {
                        back += curMeaning + Environment.NewLine;
                        if (this.sentencesList[i] != null && this.sentencesList[i].Count > 0 )
                        {
                            back += Environment.NewLine;
                            foreach (string curSentence in this.sentencesList[i])
                            {
                                back += "- " + curSentence + Environment.NewLine;
                            }
                            back += Environment.NewLine;
                        }
                        ++i;
                    }
                }

                back = StringUtil.removeRedundantNewLineCharacters(back);
                return back;
            }
        }
    }
}

