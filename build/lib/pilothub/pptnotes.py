from pilothub.pptx2content import PPTxFile


class PPTx2Notes(PPTxFile):
    def __init__(self, file_path):
        super().__init__(file_path)

    def set_skip_slides(self, skip_slides_index: list[int] = None,
                        skip_slides_layout: list[str] = None,
                        skip_slide_titles: list[str] = None):
        """
        Set the slides to skip.
        :param skip_slides_index: Index of the slides to skip.
        :param skip_slides_layout: Layout of the slides to skip.
        :param skip_slide_titles: Titles of the slides to skip.
        """
        if skip_slides_index is None:
            skip_slides_index = [1, -1]
        if skip_slides_layout is None:
            skip_slides_layout = ["Title Slide", "CoverPage", "Quote Slide", 
                                  "Agenda", "Section Header", "QuoteHead"]
        if skip_slide_titles is None:
            skip_slide_titles = []
        self.skip_slides_index = skip_slides_index
        self.skip_slides_layout = skip_slides_layout
        self.skip_slide_titles = skip_slide_titles
    
    def write_notes_to_pptx(self, output_path: str, 
                            content_client,
                            SET_SLIDE_TEXT_FOR_SKIP_SLIDES: bool = True,
                            SET_AI_TEXT_FOR_SKIP_SLIDES: bool = False,
                            AI_PROPMT_SKIP_SLIDES: str = None,
                            AI_PROMPT_DICT_SKIP_SLIDES: dict[str, str] = None,
                            DEFAULT_PROMPT_FOR_OTHER_SLIDES: str = None,
                            ):
        """
        Write notes to the PPTx file.
        :param output_path: Path to save the PPTx file.
        :param content_client: Content client to use for generating notes.
        :param SET_SLIDE_TEXT_FOR_SKIP_SLIDES: Whether to set slide text for 
                                                skip slides.
        :param SET_AI_TEXT_FOR_SKIP_SLIDES: Whether to set AI text for skip
                                            slides.
        :param AI_PROPMT_SKIP_SLIDES: Prompt to use for skip slides.
        :param AI_PROMPT_DCIT_SKIP_SLIDES: Dictionary of prompts to use for
                                            skip slides.
        :param DEFAULT_PROMPT_FOR_OTHER_SLIDES: Prompt to use for other slides.
        """
        self.content_client = content_client
        for i, slide in enumerate(self.slides):
            slide_text = self.get_slide_text(slide)
            layout = slide.slide_layout.name
            # check if skip slide
            if (i in self.skip_slides_index) or \
                (layout in self.skip_slides_layout):
                if SET_SLIDE_TEXT_FOR_SKIP_SLIDES:
                    if SET_AI_TEXT_FOR_SKIP_SLIDES:
                        if AI_PROMPT_DICT_SKIP_SLIDES is None:
                            if AI_PROPMT_SKIP_SLIDES is None:
                                raise ValueError("Please provide a prompt for AI.")
                            else:
                                prompt = AI_PROPMT_SKIP_SLIDES
                        else:
                            if layout not in AI_PROMPT_DICT_SKIP_SLIDES:
                                prompt = AI_PROPMT_SKIP_SLIDES
                            else:
                                prompt = AI_PROMPT_DICT_SKIP_SLIDES[layout]                    
                        notes_text = self.content_client.get_notes_from_text(
                            text=slide_text, prompt=prompt)
                        self.set_slide_notes(slide, notes_text)
                    else:
                        self.set_slide_notes(slide, slide_text)
                else:
                    self.erase_slide_notes(slide)
            elif slide.shapes.title and slide.shapes.title.text.lower().strip() in self.skip_slide_titles:
                self.set_slide_notes(slide, slide_text)
            else:
                prompt = DEFAULT_PROMPT_FOR_OTHER_SLIDES
                notes_text = self.content_client.get_notes_from_text(
                    text=slide_text, prompt=prompt)
                self.set_slide_notes(slide, notes_text)
        self.prs.save(output_path)
