#include <stdio.h>
#include <string.h>
#include <atlbase.h>
#include <sapi.h>
#include <comip.h>
#include <spuihelp.h>
#include <map>
#include <string>
#include <utility>
#include <vector>
#include <cstdlib>

class
CCoInitialize
{
public:
    CCoInitialize() : hr(CoInitialize(NULL)) { }
    ~CCoInitialize() {
        if (SUCCEEDED(hr))
            CoUninitialize();
    }
    operator HRESULT () const {
        return hr;
    }
private:
    HRESULT hr;
};

wchar_t *CharToWchar(const char *str);
long FileSize(const char *file_name);
HRESULT EnumerateVoices();
void ReleaseVoices();
HRESULT FileToWav(CComPtr<ISpVoice> spVoice, ISpObjectToken *voice, const char *textFile, const char *wavFile);

typedef std::map<std::string, ISpObjectToken *> VoiceMap_t;
VoiceMap_t VoiceMap;

int
main(int argc, char **argv)
{
    CCoInitialize init;
    if (FAILED(init)) {
        printf("Failed initializing COM.\n");
        return EXIT_FAILURE;
    }

    CComPtr<ISpVoice> spVoice;
    HRESULT hr = spVoice.CoCreateInstance(CLSID_SpVoice);

    if (FAILED(hr)) {
        printf("Failed creating SAPI voice.\n");
        return EXIT_FAILURE;
    }

    hr = EnumerateVoices();
    if (FAILED(hr)) {
        printf("Failed enumerating SAPI voices.\n");
        return EXIT_FAILURE;
    }

    if (argc == 2 && strcmp(argv[1], "--list-voices") == 0) {
        VoiceMap_t::const_iterator begin = VoiceMap.begin();
        VoiceMap_t::const_iterator end = VoiceMap.end();
        VoiceMap_t::const_iterator it;

        printf("The following voices are available:\n");
        for (it = begin; it != end; it++)
            printf("%s\n", it->first.c_str());

        ReleaseVoices();
        return EXIT_SUCCESS;
    }

    if (argc != 4) {
        printf("Usage: %s <voice name> <text file> <wav file>\n"
               "- or -\n"
               "Usage: %s --list-voices\n",
               argv[0], argv[0]);

        ReleaseVoices();
        return EXIT_FAILURE;
    }

    char *voice = argv[1];
    char *textFile = argv[2];
    char *wavFile = argv[3];

    VoiceMap_t::const_iterator voiceIt = VoiceMap.find(voice);
    if (voiceIt == VoiceMap.end()) {
        printf("Voice not found.\n");
    }
    else {
        hr = FileToWav(spVoice, voiceIt->second, textFile, wavFile);
    }

    ReleaseVoices();

    if (SUCCEEDED(hr))
        return EXIT_SUCCESS;
    return EXIT_FAILURE;
}

HRESULT
FileToWav(CComPtr<ISpVoice> spVoice, ISpObjectToken *voice, const char *textFile, const char *wavFile)
{
    CComPtr <ISpStream> spStream;
    CSpStreamFormat spAudioFmt;

    HRESULT hr = spVoice->SetVoice(voice);
    if (FAILED(hr)) {
        printf("Failed setting voice.\n");
        return hr;
    }

    hr = spAudioFmt.AssignFormat(SPSF_48kHz16BitStereo);
    if (FAILED(hr)) {
        printf("Failed setting audio format.\n");
        return hr;
    }
   
    hr = SPBindToFile(wavFile, SPFM_CREATE_ALWAYS, &spStream,
        &spAudioFmt.FormatId(), spAudioFmt.WaveFormatExPtr());
    if (FAILED(hr)) {
        printf("Failed binding to wav file.\n");
        return hr;
    }
	
    hr = spVoice->SetOutput(spStream, TRUE);
    if (FAILED(hr)) {
        printf("Failed setting output stream.\n");
        return hr;
    }

    //hr = spVoice->Speak(text_file, SPF_IS_FILENAME | SPF_IS_XML, NULL);
    //can't do the above as it doesn't speak utf8

    long fileSize = FileSize(textFile);
    if (fileSize < 0) {
        printf("Failed getting file size for file %s.\n", textFile);
        return -1;
    }

    wchar_t *fileContents = (wchar_t *)calloc(sizeof(wchar_t), fileSize + 1);
    if (!fileContents) {
        printf("Failed allocating memory for file %s.\n", textFile);
        return -1;
    }

    wchar_t *textFileW = CharToWchar(textFile);
    FILE *in = _wfopen(textFileW, L"rt, ccs=UTF-8");
    if (!in) {
        printf("Failed opening %s for input.\n", textFile);
        SysFreeString(textFileW);
        free(fileContents);
        return -1;
    }
    SysFreeString(textFileW);
    size_t bytes = fread(fileContents, sizeof(wchar_t), fileSize, in);
    fclose(in);

    hr = spVoice->Speak(fileContents, SPF_IS_XML, NULL);
    free(fileContents);

    if (FAILED(hr)) {
        printf("Failed speaking the contents of the file.\n");
        return hr;
    }
    spStream->Close();

    return 0;
}

wchar_t *
CharToWchar(const char *str) {
    int lenW, lenA = lstrlenA(str);
    wchar_t *unicodestr;

    lenW = MultiByteToWideChar(CP_ACP, 0, str, lenA, NULL, 0);
    if (lenW > 0) {
        unicodestr = SysAllocStringLen(0, lenW);
        MultiByteToWideChar(CP_ACP, 0, str, lenA, unicodestr, lenW);
        return unicodestr;
    }
    return NULL;
}

long
FileSize(const char *file_name)
{
    BOOL ok;
    WIN32_FILE_ATTRIBUTE_DATA fileInfo;

    if (!file_name)
        return -1;

    ok = GetFileAttributesExA(file_name, GetFileExInfoStandard, (void*)&fileInfo);
    if (!ok)
        return -1;

    return (long)fileInfo.nFileSizeLow;
}

HRESULT
EnumerateVoices()
{
    ISpObjectToken *spToken;
    CComPtr<IEnumSpObjectTokens> spEnum;

    HRESULT hrTokens = SpEnumTokens(SPCAT_VOICES, NULL, NULL, &spEnum);
    if (SUCCEEDED(hrTokens)) {
        while (spEnum->Next(1, &spToken, NULL) == S_OK) {
            CSpDynamicString spDesc;
            HRESULT hr = SpGetDescription(spToken, &spDesc);
            if (SUCCEEDED(hr)) {
                char *name = spDesc.CopyToChar();
                VoiceMap.insert(std::make_pair(name, spToken));
                CoTaskMemFree(name);
            }
            else {
                spToken->Release();
            }
        }
    }
    return hrTokens;
}

void
ReleaseVoices()
{
    VoiceMap_t::const_iterator begin = VoiceMap.begin();
    VoiceMap_t::const_iterator end = VoiceMap.end();
    VoiceMap_t::const_iterator it;

    for (it = begin; it != end; it++)
        it->second->Release();
}
