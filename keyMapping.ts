// Key mappings for Shark and Monkey libraries
export const COMPONENT_KEY_MAPPING: Record<string, Record<string, string>> = {
  'Shark': {
    'Rectangle': '3e1bf9f255bc97d64c91dea5f8308d3a174974ea',
    'Shape=Square': 'a3727e05ca34d16f76cb2e27fc31a9914ec17b6d',
    'Shape=Round': '89b3646550532710b0261b1a04c81138396895b9',
    'Tall': '78741a231050a9ae3572551f34d6c25a776c6453',
  },
  'Monkey': {
    'Rectangle': 'd25b598e077fe348e9a16efbea58c13cbaeaaa25',
    'Shape=Square': '2e78071ac513ccf3f9d7f95e8fc66787907660dd',
    'Shape=Round': '8d4fdf8e0bdccf82ac065ee1a13c0f8edf6587e0',
    'Tall': '05ebefcbebdeaec58afdd83641d229928597f999',
  },
};

export const STYLE_KEY_MAPPING: Record<string, Record<string, string>> = {
  'Shark': {
    'Background': 'd229a7e81e7a05731f1eaa36470abf5a4fae9bf2',
    'Primary': '82e40649545abfa72bf66cdd4b2f3796dd69a466',
    'HeadingColor': '19e42894199db9e7eac439f6b25f2775e0ccf651', // PAINT
    'Heading': 'ac4330274941ead0a2d0ffb00d1ea79598d22ab7', // TEXT
    'Body': '8c5340c46ff3b63418c05f0f3dafcac9e1e40b1e', // TEXT
  },
  'Monkey': {
    'Primary': 'ff2ac86fbfdde699eea044a240b1eeacf96d8a4e',
    'Background': '5c7e364fb93f9409be83e7b52d90bb2fe4663777',
    'HeadingColor': 'a190e78b22bda4b2067574705f09b42921a38ffb', // PAINT (inferred)
    'Heading': '12c7bf83a29b502cdef4352008bdd1dc87b2f93e', // TEXT
    'Body': '2f437791c4d64520626bc45e2f24a91dfc5877e9', // TEXT
  },
};

export const VARIABLE_ID_MAPPING: Record<string, Record<string, string>> = {
  'Monkey': {
    'Primary': 'VariableID:92:24',
    'Background': 'VariableID:92:25',
    'HeadingColor': 'VariableID:92:26',
  },
};

export const VARIABLE_KEY_MAPPING: Record<string, Record<string, string>> = {
  'Monkey': {
    'Primary': 'dc69fade742a1338bc34ec90e4081f924f45fbbb',
    'Background': 'e3731549dd33521e379b6720c78064011ff0c04f',
    'HeadingColor': 'bdc3548622b342867bc8976b5c780623efcf7f2a',
  },
};

export const LIBRARY_THUMBNAILS: Record<string, string> = {
  'Shark': 'https://raw.githubusercontent.com/RobertNavaille/Advanced-Library-Swap/main/Images/sharkThumbnail.png',
  'Monkey': 'https://raw.githubusercontent.com/RobertNavaille/Advanced-Library-Swap/main/Images/monkeyThumbnail.png'
};
