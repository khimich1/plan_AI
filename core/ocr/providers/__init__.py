#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from core.ocr.providers.base import OcrProvider
from core.ocr.providers.gigachat import GigaChatProvider
from core.ocr.providers.openai import OpenAIProvider

__all__ = ["OcrProvider", "OpenAIProvider", "GigaChatProvider"]
