from rutermextract import TermExtractor as TE

te = TE()

for word in te(open('text.txt', 'r').read(), strings=1):
	if len(word) > 20:
		print(word)

